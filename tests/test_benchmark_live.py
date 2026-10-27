import json
from datetime import date
import pytest
from app import auth
from app.utils import benchmark_live as live, benchmark_model as model
from app.utils.client_benchmarks import SourceError


@pytest.fixture
def isolated(monkeypatch,tmp_path):
    monkeypatch.setenv('CLIENT_BENCHMARK_CACHE_PATH',str(tmp_path/'cache.json'))
    monkeypatch.setattr(live,'_snapshot',None)
    monkeypatch.setattr(live,'_attempt',0)
    monkeypatch.setattr(live,'_error',None)
    monkeypatch.setattr(model.legacy,'london_today',lambda:date(2026,10,27))


def tabs(count=1):
    head=['Timestamp','Client','Lead company size','Lead category','Email (AMs) or Name (SDRs)']
    return {'leads':[head,*[['2026-07-10','Private client','1-10','Lead',f'{i}@example.com'] for i in range(count)]],
            'controls':[model.CONTROL_HEADERS,['Private client','2026-07','full','approved',1,0,0,0,0,0,0,0,'approved dated brief']]}


def metadata():
    return {'tabs':{},'notices':[],'eri_reference':'Not added'}


def test_refresh_change_and_allocation_change_propagate(isolated,monkeypatch):
    data=tabs()
    monkeypatch.setattr(live,'fetch_tabs',lambda:(data,metadata()))
    first=live.get_report()
    assert first['groups'][0]['qualifying_leads']==1
    data['leads']=tabs(2)['leads']
    assert live.get_report()['groups'][0]['qualifying_leads']==1  # cached for 30 minutes
    monkeypatch.setattr(live,'_attempt',0)
    assert live.get_report()['groups'][0]['qualifying_leads']==2
    data['controls'][1][4:12]=[.5,0,0,.5,0,0,0,0]
    monkeypatch.setattr(live,'_attempt',0)
    assert live.get_report()['groups'][0]['average']==4
    assert live.get_report()['groups'][1]['equivalent_months']==.5


def test_failure_retains_persisted_success_and_marks_stale(isolated,monkeypatch):
    monkeypatch.setattr(live,'fetch_tabs',lambda:(tabs(),metadata()))
    first=live.get_report()
    assert first['methodology_version']==model.METHODOLOGY_VERSION==4
    assert live.cache_path().exists()
    assert (live.cache_path().stat().st_mode & 0o777)==0o600
    monkeypatch.setattr(live,'_snapshot',None)
    monkeypatch.setattr(live,'_attempt',0)
    def fail():raise RuntimeError('PRIVATE KEY CONTENT')
    monkeypatch.setattr(live,'fetch_tabs',fail)
    stale=live.get_report()
    assert stale['stale']
    assert stale['methodology_version']==4
    assert stale['synced_at']==first['synced_at']
    assert 'PRIVATE KEY CONTENT' not in json.dumps(stale)
    assert 'audit' not in stale and 'reconciliation' not in stale
    assert 'Private client' not in json.dumps(stale)
    assert live.get_report(internal=True)['audit'][0]['client']=='private client'


def test_previous_tiers_are_not_restored_from_custom_cache_on_failed_sync(isolated,monkeypatch):
    previous=model.calculate(tabs())
    previous.update(methodology_version=3,synced_at=live.now())
    previous['groups'][0].update(headcount='1–50',bands=[0,1])
    previous['groups'][1].update(headcount='51–500',bands=[2,3])
    previous['groups'][2].update(headcount='501+',bands=[4,5,6,7])
    live.persist(previous)
    monkeypatch.setattr(live,'fetch_tabs',lambda: (_ for _ in ()).throw(SourceError('Unavailable')))

    with pytest.raises(SourceError,match='Unavailable'):
        live.get_report()
    assert live._snapshot is None
    assert json.loads(live.cache_path().read_text())['methodology_version']==3


def test_default_cache_filename_tracks_model_version(monkeypatch):
    monkeypatch.delenv('CLIENT_BENCHMARK_CACHE_PATH',raising=False)
    assert live.cache_path().name==f'unstuck-client-benchmarks-v{model.METHODOLOGY_VERSION}.json'


def test_failed_first_sync_does_not_generate_benchmarks(isolated,monkeypatch):
    monkeypatch.setattr(live,'fetch_tabs',lambda: (_ for _ in ()).throw(SourceError('Unavailable')))
    with pytest.raises(SourceError,match='Unavailable'): live.get_report()
    assert not live.cache_path().exists()


def test_internal_audit_and_template_require_admin(client,monkeypatch,isolated):
    monkeypatch.setattr(live,'fetch_tabs',lambda:(tabs(),metadata()))
    monkeypatch.delenv('AUTH_DISABLED',raising=False)
    monkeypatch.setattr(auth,'get_session_user',lambda req:{'role':'sdr','tools':['client_benchmarks'],'email':'staff@unstuck-agency.com'})
    assert client.get('/api/client-benchmarks').status_code==200
    assert client.get('/api/client-benchmarks/audit').status_code==403
    assert client.get('/api/client-benchmarks/control-template').status_code==403
    monkeypatch.setattr(auth,'get_session_user',lambda req:{'role':'admin','tools':['client_benchmarks'],'email':'admin@unstuck-agency.com'})
    audit=client.get('/api/client-benchmarks/audit')
    assert audit.status_code==200
    assert 'audit' in audit.json()
    assert client.get('/api/client-benchmarks/control-template').status_code==200


def test_source_ids_batched_read_only_and_reference_not_merged(monkeypatch):
    class Response:
        def __init__(self,value):self.value=value
        def raise_for_status(self):pass
        def json(self):return self.value
    calls=[]
    class Session:
        closed=False
        def get(self,url,params,timeout):
            calls.append((url,params))
            if live.ERI_SHEET in url:raise RuntimeError('Reference inaccessible')
            if url.endswith('batchGet'):
                return Response({'valueRanges':[{'values':[[k]]} for k in live.TABS]})
            return Response({'sheets':[{'properties':{'sheetId':gid,'title':title}} for gid,title in live.TABS.values()]})
        def close(self):self.closed=True
    session=Session()
    monkeypatch.setattr(live.source,'_is_configured',lambda:True)
    monkeypatch.setattr(live.source,'_read_session',lambda:session)
    monkeypatch.setattr(live._read_limiter,'acquire',lambda:None)
    data,meta=live.fetch_tabs()
    assert session.closed
    assert set(data)==set(live.TABS)
    assert meta['tabs']['leads']['gid']==1768027115
    assert 'unverified' in meta['eri_reference']
    assert len(calls)==3
    assert calls[1][1]['valueRenderOption']=='UNFORMATTED_VALUE'


def test_scheduler_polling_and_unconfigured_noop(monkeypatch):
    waits=[];calls=[]
    class Stop:
        def clear(self):pass
        def is_set(self):return False
        def wait(self,seconds):
            waits.append(seconds)
            return len(waits)==3
    class Thread:
        def __init__(self,target,**kwargs):self.target=target
        def start(self):self.target()
        def is_alive(self):return False
    monkeypatch.setattr(live,'_thread',None)
    monkeypatch.setattr(live,'_stop',Stop())
    monkeypatch.setattr(live.threading,'Thread',Thread)
    monkeypatch.setattr(live,'refresh',lambda:calls.append(True))
    monkeypatch.setattr(live.source,'_is_configured',lambda:False)
    live.start_scheduler()
    assert not calls
    monkeypatch.setattr(live.source,'_is_configured',lambda:True)
    live.start_scheduler()
    assert len(calls)==2
    assert waits==[10,1800,1800]


def test_old_snapshot_withholds_headline_numbers(isolated,monkeypatch):
    monkeypatch.setattr(live,'fetch_tabs',lambda:(tabs(),metadata()))
    live.get_report()
    live._snapshot['synced_at']='2020-01-01T00:00:00+00:00'
    r=live.get_report()
    assert r['stale']
    assert all(g['display']=='Data too old' for g in r['groups'])


def test_cached_display_is_updated_without_changing_underlying_rate(isolated,monkeypatch):
    monkeypatch.setattr(live,'fetch_tabs',lambda:(tabs(),metadata()))
    live.get_report()
    group=live._snapshot['groups'][1]
    group.update(average=4.19,sufficient=True,display='3–5')
    updated=live.get_report()['groups'][1]
    assert updated['display']=='~5'
    assert updated['average']==4.19
