from datetime import date
import pytest
from app.utils import benchmark_model as m
from app.utils.client_benchmarks import SourceError

HEAD = ['Timestamp','Client','Lead company size','Lead category','Email (AMs) or Name (SDRs)']
TODAY = date(2026,10,27)


def lead(client='A', when='2026-07-10', band='1-10', category='Lead', email=''):
    return [when,client,band,category,email]


def controls(*rows):
    return [m.CONTROL_HEADERS, *rows]


def control(client='A', period='2026-07', weights=None, status='full', basis='approved'):
    return [client,period,status,basis,*(weights or [1,0,0,0,0,0,0,0]),'Reviewed campaign log']


def dataset(rows, ctrls=()):
    return {'leads':[HEAD,*rows], 'controls':controls(*ctrls)}


def test_rolling_complete_months_and_no_pre_2025_baseline():
    assert m.window(TODAY,12) == [f'2025-{i:02}' for i in (10,11,12)] + [f'2026-{i:02}' for i in range(1,10)]
    assert m.window(date(2026,11,1),12)[0] == '2025-11'
    assert m.window(date(2025,4,1),18) == ['2025-01','2025-02','2025-03']


def test_exclusions_interested_identity_dedup_and_unknowns():
    records,q = m.parse_leads([HEAD,lead('Unstuck'),lead('ERI Institute'),lead('Test Client'),
        lead(email='person@example.com'),lead(when='2026-07-11',category='Interested',email='person@example.com'),
        lead('B',category='Interested'),lead('B',category='Interested'),lead('C',band='unknown'),
        lead('D',when='2026-11-01'),lead('E',category='Not interested')],TODAY)
    assert len(records)==3
    assert q['duplicates']==2
    assert q['excluded_internal_eri_test']==3
    assert q['qualifying_interested']==1
    assert q['unknown_size']==1


def test_pooled_weighted_exposure_includes_zero_and_does_not_average_client_rates():
    rows = [lead(email=f'{i}@example.com') for i in range(4)] + [lead('B')]
    rules = [control(weights=[.5,0,.3,0,.2,0,0,0]),control(period='2026-08',weights=[.5,0,.3,0,.2,0,0,0]),control('B')]
    result = m.calculate(dataset(rows,rules),TODAY)
    small = result['groups'][0]
    assert small['qualifying_leads']==5
    assert small['equivalent_months']==2
    assert small['average']==2.5
    assert small['verified_live_client_months']==3
    assert result['groups'][1]['equivalent_months']==pytest.approx(.6)
    assert result['groups'][1]['average']==0
    assert small['display']=='Insufficient data'


def test_normalise_all_eight_before_grouping_and_effective_month_overrides():
    r=m.calculate(dataset([lead(),lead(when='2026-08-10',band='1001-5000')],
        [control(weights=[40,40,10,10,10,10,10,10]),control(period='2026-08',weights=[0,0,0,0,0,1,0,0],basis='outreach')]),TODAY)
    assert r['groups'][0]['equivalent_months']==pytest.approx(80/140)
    assert r['groups'][3]['equivalent_months']==pytest.approx(30/140+1)
    assert r['groups'][3]['qualifying_leads']==1
    assert r['groups'][3]['allocation_percent']['outreach']>0


@pytest.mark.parametrize('status',['partial','inactive','unknown'])
def test_explicit_nonlive_month_never_falls_back(status):
    tabs=dataset([lead()], [control(status=status)])
    result=m.calculate(tabs,TODAY)
    assert all(g['equivalent_months']==0 for g in result['groups'])
    assert result['audit'][0]['eligible'] is False


def weight_tab():
    return [['']*23+list(m.BANDS), ['']*22+['a']+[.5,.5,0,0,0,0,0,0]]


def test_bounded_fallback_includes_quiet_month_excludes_partial_boundaries():
    tabs=dataset([lead(when='2026-06-20'),lead(when='2026-09-12')])
    tabs['weights']=weight_tab()
    r=m.calculate(tabs,TODAY)
    small=r['groups'][0]
    assert small['equivalent_months']==2
    assert small['qualifying_leads']==0
    assert small['estimated_live_client_months']==2
    assert small['allocation_percent']['estimated']==100


def test_unknown_allocation_never_inferred_from_lead_distribution():
    r=m.calculate(dataset([lead()]),TODAY)
    assert r['groups'][0]['average'] is None
    assert r['quality']['unallocated_segment_leads']==1


def test_enterprise_extends_only_when_insufficient_and_signals_low_confidence():
    r=m.calculate(dataset([lead(band='1001-5000',when='2025-05-20')],
        [control(period='2025-05',weights=[0,0,0,0,0,1,0,0])]),TODAY)
    assert len(r['groups'][0]['months'])==12
    assert len(r['groups'][3]['months'])==18
    assert r['groups'][3]['qualifying_leads']==1
    assert r['groups'][3]['display']=='Insufficient data'


def test_partial_current_month_excluded_even_with_verified_control():
    r=m.calculate(dataset([lead(when='2026-10-20')],[control(period='2026-10')]),TODAY)
    assert all(g['equivalent_months']==0 for g in r['groups'])


def test_control_duplicate_invalid_and_conflicting_weights_fail_closed():
    for rows in ([control(),control()], [control(weights=[-1,0,0,0,0,0,0,0])], [control(basis='made up')]):
        with pytest.raises(SourceError):m.parse_controls(controls(*rows))
    assert m.norm_weights(['50%',.5,0,0,0,0,0,0]) is None
    assert m.norm_weights([float('nan'),0,0,0,0,0,0,0]) is None


def test_all_boundaries_are_nonoverlapping():
    assert [i for _,_,indices in m.SEGMENTS for i in indices]==list(range(8))
    assert [m.legacy.band_index(str(n)) for n in (10,11,50,51,200,201,500,501,1000,1001)]==[0,1,1,2,2,3,3,4,4,5]


def test_reconciliation_does_not_add_report_totals_or_eri_tracker():
    tabs=dataset([lead()],[control()])
    tabs['monthly']=[['Clients','2026-07','2026-08'],['A',99,0]]
    tabs['eri_reference']=[HEAD,lead('ERI')]*10
    r=m.calculate(tabs,TODAY)
    assert r['groups'][0]['qualifying_leads']==1
    assert r['quality']['reconciliation_mismatches']==1
    assert r['reconciliation'][0]['difference']==-98


def test_dominance_and_leave_one_out_are_reported_not_removed():
    rows=[lead('A',email=f'{i}@example.com') for i in range(9)]+[lead('B'),lead('C')]
    r=m.calculate(dataset(rows,[control(c,period=p) for c in ['A','B','C'] for p in ['2026-07','2026-08']]),TODAY)
    small=r['groups'][0]
    assert small['sufficient']
    assert small['qualifying_leads']==11
    assert small['dominant_client_lead_share']==pytest.approx(9/11)
    assert small['leave_one_client_out'][0]==.5
    assert 'over half' in small['confidence']
