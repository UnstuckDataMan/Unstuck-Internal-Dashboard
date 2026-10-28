"""Append-only provider usage, priced from explicit historical configuration.

No calls means zero API spend; missing usage or rates means unknown spend.
Only Copy Studio/evaluation calls inside ``metering`` enter this ledger.
"""
from contextlib import contextmanager
from contextvars import ContextVar
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
import json
import math
import os
import uuid
from zoneinfo import ZoneInfo

from .store import StoreError

_context = ContextVar('copy_studio_cost_context', default=None)
_LONDON = ZoneInfo('Europe/London')
_MAX_TOKENS = 2**53 - 1


class CostRecordingError(StoreError):
    """A potentially charged provider call could not be durably recorded."""


def _now():
    return datetime.now(timezone.utc)


@contextmanager
def metering(store, job):
    if any(not isinstance(job.get(k), str) or not job[k].strip()
           for k in ('id', 'client_id', 'created_by')):
        raise CostRecordingError('Usage metering needs a job, client and authenticated actor.')
    if job.get('agency_id') and str(job['agency_id']) != str(store.agency_id):
        raise CostRecordingError('The metered job belongs to a different agency.')
    token = _context.set({'store': store, 'job': dict(job), 'calls': {}})
    try:
        yield
    finally:
        _context.reset(token)


def _count(value):
    return value if type(value) is int and 0 <= value <= _MAX_TOKENS else None


def _usage(provider, raw):
    result = dict(input_tokens=None, output_tokens=None, cached_input_tokens=None,
                  cache_creation_input_tokens=None, uncached_input_tokens=None,
                  usage_known=False)
    if not isinstance(raw, dict):
        return result
    initial, output = _count(raw.get('input_tokens')), _count(raw.get('output_tokens'))
    if provider == 'openai':
        details = raw.get('input_tokens_details', {})
        cached = _count(details.get('cached_tokens', 0)) if isinstance(details, dict) else None
        created = 0
        total = initial
        uncached = initial - cached if initial is not None and cached is not None else None
    elif provider == 'anthropic':
        # Anthropic reports uncached input separately from cache reads/writes.
        cached = _count(raw.get('cache_read_input_tokens', 0))
        created = _count(raw.get('cache_creation_input_tokens', 0))
        total = initial + cached + created if None not in (initial, cached, created) else None
        uncached = initial
    else:
        return result
    result.update(input_tokens=total, output_tokens=output, cached_input_tokens=cached,
                  cache_creation_input_tokens=created, uncached_input_tokens=uncached)
    result['usage_known'] = (None not in (total, output, cached, created, uncached)
                             and all(0 <= n <= _MAX_TOKENS for n in (total, output, cached, created, uncached)))
    return result


def _pricing(provider, model, usage):
    """Rates are USD per million tokens, never inferred from a model name."""
    try:
        configured = json.loads(os.environ.get('COPY_STUDIO_PRICING_JSON', '{}'))
        rates = configured.get(f'{provider}:{model}') if isinstance(configured, dict) else None
        if not isinstance(rates, dict):
            return None, None, 'No pricing is configured for this provider and model.'
        parsed = {}
        for key, value in rates.items():
            if key not in ('input', 'output', 'cached_input', 'cache_creation_input'):
                raise ValueError
            if type(value) not in (int, float) or not math.isfinite(value) or value < 0:
                raise ValueError
            parsed[key] = Decimal(str(value))
        if not {'input', 'output'} <= parsed.keys():
            raise ValueError
        if usage['cached_input_tokens'] and 'cached_input' not in parsed:
            return None, None, 'Cached-input pricing is not configured.'
        if usage['cache_creation_input_tokens'] and 'cache_creation_input' not in parsed:
            return None, None, 'Cache-creation pricing is not configured.'
        amount = (usage['uncached_input_tokens'] * parsed['input'] + usage['output_tokens'] * parsed['output']
                  + usage['cached_input_tokens'] * parsed.get('cached_input', Decimal(0))
                  + usage['cache_creation_input_tokens'] * parsed.get('cache_creation_input', Decimal(0))) / Decimal(1_000_000)
        if not math.isfinite(float(amount)):
            raise ValueError
        return {k: str(v) for k, v in parsed.items()}, float(amount), ''
    except (ValueError, TypeError, InvalidOperation, OverflowError):
        return None, None, 'Configured pricing is invalid.'


def _save(context, call_id, content, parent_id=None):
    try:
        record = context['store'].save('usage', context['job']['client_id'], 'api:'+call_id,
                                       content, context['job']['created_by'], parent_id=parent_id)
        if not record:
            raise StoreError('No saved usage record was returned.')
        return record
    except Exception as exc:
        raise CostRecordingError('API usage could not be saved. A started call remains unpriced until '
                                 'reconciled; check Copy Studio storage before retrying.') from exc


def start_call(provider, model):
    """Persist an unknown attempt before network I/O; failure prevents the call."""
    context = _context.get()
    if context is None:
        return None
    if not isinstance(provider, str) or not isinstance(model, str) or not provider or not model:
        raise CostRecordingError('Usage metering requires the provider and model.')
    call_id = str(uuid.uuid4())
    job, store = context['job'], context['store']
    content = {
        'schema_version': 1, 'call_id': call_id, 'phase': 'started',
        'provider': provider, 'model': model, 'request_id': None,
        'response_status': 'started', 'job_id': job['id'], 'job_kind': job.get('kind', ''),
        'agency_id': str(store.agency_id), 'client_id': job['client_id'],
        'occurred_at': _now().astimezone(timezone.utc).isoformat(),
        **_usage(provider, None), 'pricing_per_million_usd': None,
        'estimated_usd': None, 'cost_known': False, 'currency': 'USD',
        'pricing_note': 'API attempt started; final usage has not been recorded.',
    }
    record = _save(context, call_id, content)
    context['calls'][call_id] = {'started': record, 'final': None}
    return call_id


def record_usage(provider, model, usage, request_id=None, status='received', call_id=None):
    """Record an API response before output parsing; no-op outside a metered job.

    Start the attempt with ``start_call`` before network I/O. Pass the provider's
    response/request ID as metadata where available. Transport errors or
    responses without usage are recorded as unknown, never as free calls.
    """
    context = _context.get()
    if context is None:
        return None
    call_id = call_id or start_call(provider, model)
    call = context['calls'].get(call_id)
    if not call or (call['started']['content']['provider'], call['started']['content']['model']) != (provider, model):
        raise CostRecordingError('The API usage does not match this metered attempt.')
    if call['final'] is not None:
        return call['final']
    normalized = _usage(provider, usage)
    prices, amount, pricing_error = (_pricing(provider, model, normalized) if normalized['usage_known']
                                     else (None, None, 'The provider did not return valid token usage.'))
    content = {
        **call['started']['content'], 'phase': 'recorded',
        'request_id': str(request_id) if request_id is not None else None,
        'response_status': str(status)[:100],
        'recorded_at': _now().astimezone(timezone.utc).isoformat(),
        **normalized, 'pricing_per_million_usd': prices,
        'estimated_usd': amount, 'cost_known': amount is not None,
        'pricing_note': pricing_error, 'currency': 'USD',
    }
    record = _save(context, call_id, content, parent_id=call['started']['id'])
    call['final'] = record
    return record


def _stamp(value):
    try:
        parsed = datetime.fromisoformat(str(value).replace('Z', '+00:00'))
        return parsed.astimezone(timezone.utc) if parsed.tzinfo is not None else None
    except (ValueError, TypeError):
        return None


def _bucket(rows):
    known = Decimal(0)
    known_calls = unknown_usage = unpriced = 0
    for row in rows:
        usage = row['content']
        value = usage.get('estimated_usd')
        valid = (usage.get('usage_known') is True and usage.get('cost_known') is True
                 and type(value) in (int, float) and math.isfinite(value) and value >= 0)
        if valid:
            known += Decimal(str(value))
            known_calls += 1
        elif usage.get('usage_known') is not True:
            unknown_usage += 1
        else:
            unpriced += 1
    complete = not (unknown_usage or unpriced)
    subtotal = float(known)
    if not math.isfinite(subtotal):
        raise StoreError('Recorded cost totals exceed the supported numeric range.')
    return {'amount': subtotal if complete else None, 'known_cost': subtotal,
            'total_calls': len(rows), 'known_calls': known_calls,
            'unpriced_calls': unpriced, 'unknown_usage_calls': unknown_usage,
            'complete': complete}


def summarize(store, client_ids, now=None):
    """All recorded tool usage for explicitly permitted clients; no DB-wide scan."""
    at = now or _now()
    if at.tzinfo is None:
        raise ValueError('Cost reporting requires a timezone-aware time.')
    local = at.astimezone(_LONDON)
    seen = {}
    for client_id in sorted(set(client_ids)):
        for row in store.iter_records(client_id, kind='usage'):
            if row.get('agency_id') != store.agency_id or row.get('client_id') != client_id:
                continue
            content = row.get('content', {})
            when = _stamp(content.get('occurred_at'))
            if when is None:
                raise StoreError('A usage record has an invalid timestamp; cost totals are unavailable.')
            if when > at.astimezone(timezone.utc):
                continue
            key = (client_id, content.get('call_id') or row['id'])
            # Final usage supersedes the pre-request unknown marker. A distinct
            # retry is another attempt even if a provider reuses a request ID.
            rank = (content.get('phase') == 'recorded', row.get('created_at', ''), row['id'])
            previous = seen.get(key)
            previous_rank = ((previous[1]['content'].get('phase') == 'recorded', previous[1].get('created_at', ''), previous[1]['id'])
                             if previous else None)
            if previous_rank is None or rank > previous_rank:
                seen[key] = (when, row)
    entries = list(seen.values())
    today = [row for when, row in entries if when.astimezone(_LONDON).date() == local.date()]
    month = [row for when, row in entries if (when.astimezone(_LONDON).year, when.astimezone(_LONDON).month) == (local.year, local.month)]
    return {'currency': 'USD', 'timezone': 'Europe/London', 'scope': 'copy_studio',
            'as_of': at.astimezone(timezone.utc).isoformat(), 'today': _bucket(today),
            'month': _bucket(month), 'all_time': _bucket([row for _, row in entries])}
