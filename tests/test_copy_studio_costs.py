"""Metering is scoped, append-only, priced explicitly and honest about gaps."""
from datetime import datetime, timezone
import json

import pytest

from app.utils.copy_studio import costs
from app.utils.copy_studio.store import SQLiteStore, StoreError

AGENCY = '11111111-1111-1111-1111-111111111111'
NOW = datetime(2026, 10, 28, 12, tzinfo=timezone.utc)
JOB = {'id': 'job-a', 'agency_id': AGENCY, 'client_id': 'client-a',
       'created_by': 'staff@example.test', 'kind': 'generate'}


@pytest.fixture
def db(tmp_path, monkeypatch):
    monkeypatch.delenv('COPY_STUDIO_PRICING_JSON', raising=False)
    monkeypatch.setattr(costs, '_now', lambda: NOW)
    return SQLiteStore(str(tmp_path / 'costs.sqlite'), AGENCY)


def prices(monkeypatch, **rates):
    monkeypatch.setenv('COPY_STUDIO_PRICING_JSON', json.dumps(rates))


def call(db, *, provider='openai', model='test-model', usage=None, request_id='response-a',
         job=None, status='received'):
    with costs.metering(db, job or JOB):
        identifier = costs.start_call(provider, model)
        return costs.record_usage(provider, model, usage, request_id=request_id,
                                  status=status, call_id=identifier)


def test_no_calls_and_keyless_assembly_are_zero_without_ledger_records(db):
    assert costs.start_call('openai', 'test-model') is None
    assert costs.record_usage('openai', 'test-model', {}) is None
    with costs.metering(db, JOB):
        pass  # Approved-template assembly makes no provider call.
    summary = costs.summarize(db, ['client-a'], NOW)
    assert summary['currency'] == 'USD'
    assert summary['timezone'] == 'Europe/London'
    for key in ('today', 'month', 'all_time'):
        assert summary[key]['amount'] == 0
        assert summary[key]['total_calls'] == 0
        assert summary[key]['complete']
    assert db.list('client-a', kind='usage') == []


def test_openai_cached_tokens_and_prices_are_persisted_not_repriced(db, monkeypatch):
    prices(monkeypatch, **{'openai:test-model': {'input': 2, 'output': 10, 'cached_input': .5}})
    record = call(db, usage={'input_tokens': 1_000_000, 'output_tokens': 100_000,
                            'input_tokens_details': {'cached_tokens': 250_000}})
    content = record['content']
    assert content['input_tokens'] == 1_000_000
    assert content['uncached_input_tokens'] == 750_000
    assert content['estimated_usd'] == pytest.approx(2.625)
    assert content['pricing_per_million_usd'] == {'input': '2', 'output': '10', 'cached_input': '0.5'}
    assert content['request_id'] == 'response-a'
    assert record['agency_id'] == AGENCY and record['client_id'] == 'client-a'
    assert record['created_by'] == JOB['created_by']
    assert len(db.list('client-a', kind='usage')) == 2  # Started + final immutable versions.
    prices(monkeypatch, **{'openai:test-model': {'input': 200, 'output': 1000, 'cached_input': 50}})
    bucket = costs.summarize(db, ['client-a'], NOW)['today']
    assert bucket['amount'] == pytest.approx(2.625)
    assert bucket['known_calls'] == bucket['total_calls'] == 1


def test_anthropic_counts_cache_reads_and_creation_separately(db, monkeypatch):
    prices(monkeypatch, **{'anthropic:test-model': {'input': 3, 'output': 15,
                                                  'cached_input': .3, 'cache_creation_input': 3.75}})
    record = call(db, provider='anthropic', usage={'input_tokens': 1_000_000, 'output_tokens': 100_000,
                  'cache_read_input_tokens': 500_000, 'cache_creation_input_tokens': 250_000})
    assert record['content']['input_tokens'] == 1_750_000
    assert record['content']['uncached_input_tokens'] == 1_000_000
    assert record['content']['estimated_usd'] == pytest.approx(5.5875)


@pytest.mark.parametrize('configured', [None, '{}', '{"openai:another-model":{"input":1,"output":2}}'])
def test_verified_sonnet_default_records_prices_and_dated_source(db, monkeypatch, configured):
    if configured is not None:
        monkeypatch.setenv('COPY_STUDIO_PRICING_JSON', configured)
    content = call(db, provider='anthropic', model='claude-sonnet-5-5', usage={
        'input_tokens': 1_000_000, 'output_tokens': 100_000,
        'cache_read_input_tokens': 500_000})['content']
    assert content['estimated_usd'] == pytest.approx(3.05)
    assert content['pricing_per_million_usd'] == {'input': '2', 'output': '10', 'cached_input': '0.1'}
    assert '2026-10-28' in content['pricing_note']
    assert 'https://platform.claude.com/docs/en/about-claude/pricing' in content['pricing_note']
    prices(monkeypatch, **{'anthropic:claude-sonnet-5-5': {'input': 200, 'output': 1000}})
    assert costs.summarize(db, ['client-a'], NOW)['all_time']['amount'] == pytest.approx(3.05)


def test_explicit_sonnet_configuration_overrides_verified_default(db, monkeypatch):
    prices(monkeypatch, **{'anthropic:claude-sonnet-5-5': {'input': 4, 'output': 20}})
    content = call(db, provider='anthropic', model='claude-sonnet-5-5', usage={
        'input_tokens': 1_000_000, 'output_tokens': 100_000})['content']
    assert content['estimated_usd'] == 6
    assert content['pricing_per_million_usd'] == {'input': '4', 'output': '20'}
    assert content['pricing_note'] == ''


@pytest.mark.parametrize('configured', [
    'not-json', '[]', 'null', '{"anthropic:claude-sonnet-5-5":null}',
    '{"anthropic:claude-sonnet-5-5":{}}',
    '{"anthropic:claude-sonnet-5-5":{"input":-1,"output":10}}',
])
def test_invalid_sonnet_configuration_never_falls_back_to_verified_prices(db, monkeypatch, configured):
    monkeypatch.setenv('COPY_STUDIO_PRICING_JSON', configured)
    content = call(db, provider='anthropic', model='claude-sonnet-5-5', usage={
        'input_tokens': 1_000_000, 'output_tokens': 0})['content']
    assert content['estimated_usd'] is None
    assert not content['cost_known']
    assert 'Default standard pricing' not in content['pricing_note']


@pytest.mark.parametrize('provider,model', [
    ('anthropic', 'claude-sonnet-5-5-latest'), ('anthropic', 'claude-sonnet-5'),
    ('openai', 'claude-sonnet-5-5'),
])
def test_verified_default_never_guesses_other_models_or_providers(db, provider, model):
    content = call(db, provider=provider, model=model, usage={
        'input_tokens': 1_000_000, 'output_tokens': 0})['content']
    assert content['estimated_usd'] is None
    assert costs.summarize(db, ['client-a'], NOW)['today']['unpriced_calls'] == 1


def test_verified_sonnet_default_leaves_unknown_duration_cache_writes_unpriced(db):
    content = call(db, provider='anthropic', model='claude-sonnet-5-5', usage={
        'input_tokens': 100, 'output_tokens': 10, 'cache_creation_input_tokens': 500})['content']
    assert content['usage_known'] and not content['cost_known']
    assert content['estimated_usd'] is None
    assert 'Cache-creation pricing is not configured' in content['pricing_note']
    assert '2026-10-28' in content['pricing_note']
    assert 'https://platform.claude.com/docs/en/about-claude/pricing' in content['pricing_note']
    assert costs.summarize(db, ['client-a'], NOW)['all_time']['unpriced_calls'] == 1


@pytest.mark.parametrize('raw', [None, {}, {'input_tokens': 5}, {'input_tokens': -1, 'output_tokens': 2},
    {'input_tokens': True, 'output_tokens': 2}, {'input_tokens': 1.5, 'output_tokens': 2},
    {'input_tokens': float('inf'), 'output_tokens': 2}, {'input_tokens': 2**53, 'output_tokens': 2},
    {'input_tokens': 5, 'output_tokens': 2, 'input_tokens_details': {'cached_tokens': 6}},
    {'input_tokens': 5, 'output_tokens': 2, 'input_tokens_details': {'cached_tokens': -1}}])
def test_missing_or_invalid_usage_remains_unknown(db, monkeypatch, raw):
    prices(monkeypatch, **{'openai:test-model': {'input': 1, 'output': 1, 'cached_input': 1}})
    record = call(db, usage=raw)
    assert not record['content']['usage_known']
    assert record['content']['estimated_usd'] is None
    bucket = costs.summarize(db, ['client-a'], NOW)['all_time']
    assert bucket['amount'] is None and bucket['known_cost'] == 0
    assert bucket['unknown_usage_calls'] == bucket['total_calls'] == 1
    assert not bucket['complete']


@pytest.mark.parametrize('rates', [{}, {'input': 1}, {'input': -1, 'output': 1},
    {'input': float('nan'), 'output': 1}, {'input': float('inf'), 'output': 1},
    {'input': True, 'output': 1}, {'input': '1', 'output': 1},
    {'input': 1, 'output': 1, 'invented': 2}])
def test_invalid_or_absent_rates_do_not_invent_prices(db, monkeypatch, rates):
    prices(monkeypatch, **{'openai:test-model': rates})
    call(db, usage={'input_tokens': 1_000_000, 'output_tokens': 0})
    bucket = costs.summarize(db, ['client-a'], NOW)['all_time']
    assert bucket['amount'] is None
    assert bucket['unpriced_calls'] == 1 and bucket['unknown_usage_calls'] == 0


@pytest.mark.parametrize('provider,usage', [
    ('openai', {'input_tokens': 100, 'output_tokens': 0, 'input_tokens_details': {'cached_tokens': 10}}),
    ('anthropic', {'input_tokens': 100, 'output_tokens': 0, 'cache_creation_input_tokens': 10}),
])
def test_used_cache_pricing_cannot_silently_fall_back_to_full_input_price(db, monkeypatch, provider, usage):
    prices(monkeypatch, **{provider+':test-model': {'input': 1, 'output': 1}})
    call(db, provider=provider, usage=usage)
    assert costs.summarize(db, ['client-a'], NOW)['today']['unpriced_calls'] == 1


def test_mixed_priced_unpriced_and_unknown_calls_show_only_known_subtotal(db, monkeypatch):
    prices(monkeypatch, **{'openai:test-model': {'input': 1, 'output': 1}})
    call(db, usage={'input_tokens': 1_000_000, 'output_tokens': 0})
    call(db, model='unpriced-model', usage={'input_tokens': 1_000_000, 'output_tokens': 0})
    call(db, usage=None, status='transport_unknown')
    bucket = costs.summarize(db, ['client-a'], NOW)['today']
    assert bucket == {'amount': None, 'known_cost': 1.0, 'total_calls': 3, 'known_calls': 1,
                      'unpriced_calls': 1, 'unknown_usage_calls': 1, 'complete': False}


def test_crash_after_start_is_visible_and_final_record_collapses_one_attempt(db, monkeypatch):
    prices(monkeypatch, **{'openai:test-model': {'input': 1, 'output': 1}})
    with costs.metering(db, JOB):
        crashed = costs.start_call('openai', 'test-model')
        bucket = costs.summarize(db, ['client-a'], NOW)['all_time']
        assert bucket['amount'] is None and bucket['unknown_usage_calls'] == 1
        final = costs.record_usage('openai', 'test-model', {'input_tokens': 1_000_000, 'output_tokens': 0},
                                   call_id=crashed, status='output_parse_failed')
        repeated = costs.record_usage('openai', 'test-model', {}, call_id=crashed)
        assert repeated == final
    bucket = costs.summarize(db, ['client-a'], NOW)['all_time']
    assert bucket['amount'] == 1 and bucket['total_calls'] == 1
    assert len(db.list('client-a', kind='usage')) == 2


def test_distinct_attempts_with_same_provider_id_are_counted_separately(db, monkeypatch):
    prices(monkeypatch, **{'openai:test-model': {'input': 1, 'output': 1}})
    for _ in range(2):
        call(db, usage={'input_tokens': 1_000_000, 'output_tokens': 0}, request_id='same-provider-id')
    assert costs.summarize(db, ['client-a'], NOW)['all_time']['amount'] == 2


def test_failed_start_prevents_call_and_failed_final_preserves_unknown_marker(db, monkeypatch):
    save = db.save
    def broken(*args, **kwargs):
        raise StoreError('storage offline')
    with costs.metering(db, JOB):
        monkeypatch.setattr(db, 'save', broken)
        with pytest.raises(costs.CostRecordingError, match='usage could not be saved'):
            costs.start_call('openai', 'test-model')
        monkeypatch.setattr(db, 'save', save)
        call_id = costs.start_call('openai', 'test-model')
        monkeypatch.setattr(db, 'save', broken)
        with pytest.raises(costs.CostRecordingError):
            costs.record_usage('openai', 'test-model', {'input_tokens': 100, 'output_tokens': 5}, call_id=call_id)
    bucket = costs.summarize(db, ['client-a'], NOW)['all_time']
    assert bucket['unknown_usage_calls'] == bucket['total_calls'] == 1
    assert bucket['amount'] is None


def test_london_month_boundary_during_bst(db, monkeypatch):
    prices(monkeypatch, **{'openai:test-model': {'input': 1, 'output': 0}})
    for stamp in ['2026-02-15T10:00:00+00:00', '2026-03-31T22:59:59+00:00',
                  '2026-03-31T23:00:00+00:00', '2026-04-01T00:20:00+00:00']:
        monkeypatch.setattr(costs, '_now', lambda s=stamp: datetime.fromisoformat(s))
        call(db, usage={'input_tokens': 1_000_000, 'output_tokens': 0})
    result = costs.summarize(db, ['client-a'], datetime.fromisoformat('2026-04-01T00:30:00+00:00'))
    assert result['today']['amount'] == result['month']['amount'] == 2
    assert result['all_time']['amount'] == 4


def test_london_fall_back_day_includes_both_repeated_hours(db, monkeypatch):
    prices(monkeypatch, **{'openai:test-model': {'input': 1, 'output': 0}})
    for stamp in ['2026-10-24T22:59:59+00:00', '2026-10-24T23:00:00+00:00',
                  '2026-10-25T00:30:00+00:00', '2026-10-25T01:30:00+00:00', '2026-10-25T23:30:00+00:00']:
        monkeypatch.setattr(costs, '_now', lambda s=stamp: datetime.fromisoformat(s))
        call(db, usage={'input_tokens': 1_000_000, 'output_tokens': 0})
    result = costs.summarize(db, ['client-a'], datetime.fromisoformat('2026-10-25T23:59:59+00:00'))
    assert result['today']['amount'] == 4
    assert result['month']['amount'] == 5


def test_client_agency_scope_and_context_cleanup(db, monkeypatch):
    prices(monkeypatch, **{'openai:test-model': {'input': 1, 'output': 0}})
    call(db, usage={'input_tokens': 1_000_000, 'output_tokens': 0})
    call(db, job={**JOB, 'client_id': 'client-b'}, usage={'input_tokens': 2_000_000, 'output_tokens': 0})
    assert costs.summarize(db, ['client-a', 'client-a'], NOW)['all_time']['amount'] == 1
    assert costs.summarize(db, ['client-b'], NOW)['all_time']['amount'] == 2
    other = SQLiteStore(db.path, '22222222-2222-2222-2222-222222222222')
    assert costs.summarize(other, ['client-a'], NOW)['all_time']['amount'] == 0
    with pytest.raises(costs.CostRecordingError, match='different agency'):
        with costs.metering(other, JOB):
            pass
    with pytest.raises(ValueError):
        with costs.metering(db, JOB):
            raise ValueError('interrupt')
    assert costs.start_call('openai', 'test-model') is None


def test_invalid_dates_and_naive_report_time_do_not_silently_drop_usage(db):
    with pytest.raises(ValueError, match='timezone-aware'):
        costs.summarize(db, ['client-a'], datetime(2026,10,28))
    db.save('usage', 'client-a', 'invalid', {'occurred_at': 'not-a-date'}, JOB['created_by'])
    with pytest.raises(StoreError, match='timestamp'):
        costs.summarize(db, ['client-a'], NOW)


def test_provider_output_parse_failure_still_records_reported_usage(db, monkeypatch):
    from app.utils.copy_studio import providers
    prices(monkeypatch, **{'openai:test-model': {'input': 1, 'output': 2}})
    class Response:
        status_code = 200
        headers = {}
        def json(self):
            return {'id': 'provider-response', 'status': 'completed',
                    'usage': {'input_tokens': 100, 'output_tokens': 50},
                    'output': [{'content': [{'type': 'output_text', 'text': 'invalid JSON'}]}]}
    monkeypatch.setattr(providers.requests, 'post', lambda *args, **kwargs: Response())
    with costs.metering(db, JOB), pytest.raises(providers.ProviderError):
        providers.OpenAIProvider('test-model', 'fake-key').complete_json(
            'system', 'user', schema={'type': 'object'}, schema_name='test')
    bucket = costs.summarize(db, ['client-a'], NOW)['today']
    assert bucket['known_calls'] == bucket['total_calls'] == 1
    assert bucket['amount'] == pytest.approx(.0002)


def test_provider_does_not_send_when_start_marker_cannot_be_saved(db, monkeypatch):
    from app.utils.copy_studio import providers
    calls = []
    monkeypatch.setattr(providers.requests, 'post', lambda *args, **kwargs: calls.append(True))
    monkeypatch.setattr(db, 'save', lambda *args, **kwargs: (_ for _ in ()).throw(StoreError('offline')))
    with costs.metering(db, JOB), pytest.raises(costs.CostRecordingError):
        providers.OpenAIProvider('test-model', 'fake-key').complete_json(
            'system', 'user', schema={'type': 'object'}, schema_name='test')
    assert calls == []
