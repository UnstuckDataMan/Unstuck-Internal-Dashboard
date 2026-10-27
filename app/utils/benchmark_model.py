"""Auditable pooled leads / allocated live client-month exposure. No fitted targets."""
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
import math
import re

from app.utils import client_benchmarks as legacy

BANDS = legacy.BANDS
SEGMENTS = [('Small businesses', '1–50', (0, 1)),
            ('Mid-market', '51–500', (2, 3)),
            ('Enterprise', '501+', (4, 5, 6, 7))]
MIN_CLIENTS = 3
MIN_EXPOSURE = 6
EXCLUDED = {'unstuck', 'unstuck agency', 'eri', 'eri institute', 'economic research institute', 'eri economic research institute'}
CONTROL_HEADERS = ['Client', 'Month', 'Activity status', 'Allocation source', *BANDS, 'Evidence']


def normal(value):
    return legacy._normal(value)


def excluded(client):
    return client in EXCLUDED or bool(re.match(r'^(test|testing|demo|sample)(\s|$)', client))


def day(value):
    if isinstance(value, (datetime, date)):
        return value.date() if isinstance(value, datetime) else value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 30000 < value < 100000:
            return (datetime(1899, 12, 30) + timedelta(days=value)).date()
        return None
    return legacy.parse_date(value)


def month(value):
    d = day(value)
    if d:
        return d.strftime('%Y-%m')
    text = str(value or '').strip()
    if re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])', text):
        return text
    for fmt in ('%b %Y', '%B %Y', '%b-%y', '%b-%Y', '%m/%Y'):
        try:
            return datetime.strptime(text, fmt).strftime('%Y-%m')
        except ValueError:
            pass
    return None


def window(today, length):
    end = today.year * 12 + today.month - 2
    # Pre-2025 categorisation is not used as an equivalent baseline.
    return [legacy.month_string(i) for i in range(max(2025 * 12, end - length + 1), end + 1)]


def norm_weights(values):
    if len(values) != 8:
        return None
    try:
        nums = [float(str(v).strip().rstrip('%')) if v not in ('', None) else 0.0 for v in values]
        if any(not math.isfinite(v) or v < 0 for v in nums) or sum(nums) <= 0:
            return None
        # Mixed percent strings and numeric fractions have ambiguous units.
        supplied = [v for v in values if v not in ('', None, 0, '0')]
        if any(str(v).endswith('%') for v in supplied) and not all(str(v).endswith('%') for v in supplied):
            return None
        return [v / sum(nums) for v in nums]
    except (ValueError, TypeError):
        return None


def parse_leads(values, today):
    required = ('timestamp', 'client', 'lead company size', 'lead category')
    header = next(((i, [normal(c) for c in row]) for i, row in enumerate(values[:10])
                   if all(x in [normal(c) for c in row] for x in required)), None)
    if not header:
        raise legacy.SourceError('Form Responses is missing its required lead columns.')
    i, names = header
    if any(names.count(k) != 1 for k in required):
        raise legacy.SourceError('Form Responses has duplicate required headers.')
    records, quality, identities, exact = [], Counter(), {}, set()
    for row in values[i+1:]:
        if not any(c not in ('', None) for c in row):
            continue
        get = lambda k: row[names.index(k)] if k in names and names.index(k) < len(row) else ''
        client, when, category = normal(get('client')), day(get('timestamp')), normal(get('lead category'))
        if excluded(client):
            quality['excluded_internal_eri_test'] += 1
            continue
        if not client or not when or when > today:
            quality['invalid_or_future'] += 1
            continue
        if category not in ('lead', 'interested'):
            quality['excluded_category'] += 1
            continue
        signature = tuple(normal(c) for c in row)
        if signature in exact:
            quality['duplicates'] += 1
            continue
        exact.add(signature)
        email = normal(get('email (ams) or name (sdrs)'))
        # Only a real email establishes lead identity; names/companies alone do not.
        identity = (client, email) if re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+', email) else None
        band = legacy.band_index(get('lead company size'))
        item = {'client': client, 'date': when.isoformat(), 'month': when.strftime('%Y-%m'), 'band': band, 'category': category}
        if identity and identity in identities:
            quality['duplicates'] += 1
            previous = identities[identity]
            if when.isoformat() < records[previous]['date']:
                records[previous] = item
            continue
        if identity:
            identities[identity] = len(records)
        else:
            quality['without_stable_lead_identity'] += 1
        records.append(item)
    quality['unknown_size'] = sum(r['band'] is None for r in records)
    quality['qualifying_lead'] = sum(r['category'] == 'lead' for r in records)
    quality['qualifying_interested'] = sum(r['category'] == 'interested' for r in records)
    return records, dict(quality)


def parse_weights(values):
    """X:AE only when its eight headers match; never use lead-outcome columns."""
    header = next((i for i, row in enumerate(values[:10])
                   if len(row) >= 31 and [legacy.band_index(v) for v in row[23:31]] == list(range(8))), None)
    if header is None:
        return {}, ['Targeting weights X:AE could not be verified; affected allocations remain unknown.']
    result, duplicates, invalid = {}, set(), 0
    for row in values[header+1:]:
        client = normal(row[22]) if len(row) > 22 else ''
        if not client or excluded(client):
            continue
        weights = norm_weights((list(row[23:31]) + [None] * 8)[:8])
        if client in result or client in duplicates:
            duplicates.add(client)
            result.pop(client, None)
        elif weights:
            result[client] = weights
        else:
            invalid += 1
    return result, [f'{invalid} invalid weight rows; {len(duplicates)} duplicate clients excluded from weight fallback.']


def parse_controls(values):
    if not values:
        return {}
    names = [normal(v) for v in values[0]]
    expected = [normal(v) for v in CONTROL_HEADERS]
    # Accept hyphen/spacing variants of the eight band labels.
    if len(names) != 13 or names[:4] != expected[:4] or names[-1] != 'evidence' or [legacy.band_index(v) for v in values[0][4:12]] != list(range(8)):
        raise legacy.SourceError('Benchmark campaign months headers must match the downloadable template.')
    controls = {}
    for index, raw in enumerate(values[1:], 2):
        if not any(v not in ('', None) for v in raw):
            continue
        row = (list(raw) + [''] * 13)[:13]
        client, period, status, basis = normal(row[0]), month(row[1]), normal(row[2]), normal(row[3])
        if excluded(client):
            continue
        key = (client, period)
        evidence = str(row[12] or '').strip()
        if not client or not period or status not in ('full', 'partial', 'inactive', 'unknown') or not evidence or key in controls:
            raise legacy.SourceError(f'Invalid or duplicate campaign-month control at row {index}. Supply activity status and evidence.')
        weights = norm_weights(row[4:12])
        if status == 'full' and (basis not in ('outreach', 'approved', 'estimated', 'unknown') or (basis != 'unknown' and weights is None)):
            raise legacy.SourceError(f'Invalid allocation at campaign-month row {index}.')
        controls[key] = {'status': status, 'basis': basis, 'weights': weights if basis != 'unknown' else None, 'evidence': evidence}
    return controls


def parse_monthly(values):
    header = next(((i, [(j, month(v)) for j, v in enumerate(row) if month(v)])
                   for i, row in enumerate(values[:10]) if sum(month(v) is not None for v in row) >= 2), None)
    if not header:
        return {}, ['Monthly leads headers could not be reconciled.']
    i, columns = header
    result, duplicates = {}, set()
    for row in values[i+1:]:
        client = normal(row[0]) if row else ''
        if not client or excluded(client):
            continue
        for j, period in columns:
            value = row[j] if j < len(row) else None
            try:
                count = float(value)
            except (ValueError, TypeError):
                continue
            if not math.isfinite(count) or count < 0 or not count.is_integer():
                continue
            key = (client, period)
            if key in result or key in duplicates:
                duplicates.add(key)
                result.pop(key, None)
            else:
                result[key] = int(count)
    return result, [f'{len(duplicates)} duplicated client/month totals omitted from reconciliation.']


def calculate(tabs, today=None):
    today = today or legacy.london_today()
    records, quality = parse_leads(tabs['leads'], today)
    weights, warnings = parse_weights(tabs.get('weights', []))
    controls = parse_controls(tabs.get('controls', []))
    reported, monthly_warnings = parse_monthly(tabs.get('monthly', []))
    warnings += monthly_warnings
    base = window(today, 6)
    if not base:
        raise legacy.SourceError('No complete post-2024 reporting months are available.')
    by_month, dates, monthly_counts = defaultdict(list), defaultdict(list), Counter()
    for r in records:
        by_month[r['client'], r['month']].append(r)
        dates[r['client']].append(date.fromisoformat(r['date']))
        monthly_counts[r['client'], r['month']] += 1
    # Reconciliation is evidence only; it never adds a second copy of a lead.
    reconciliation = [{'client': c, 'month': m, 'raw_qualifying_leads': monthly_counts[c,m],
                       'reported_leads': count, 'difference': monthly_counts[c,m]-count}
                      for (c,m), count in sorted(reported.items()) if m in base]
    clients = set(dates) | {c for c, m in controls}
    audit = []
    for client in sorted(clients):
        for period in base:
            entry = controls.get((client, period))
            rows = by_month[client, period]
            if entry:
                status, allocation = entry['status'], entry['weights']
                basis, activity = entry['basis'], 'verified' if status == 'full' else entry['status']
            else:
                start = date.fromisoformat(period+'-01')
                index = start.year*12+start.month-1
                end = date.fromisoformat(legacy.month_string(index+1)+'-01')-timedelta(days=1)
                # Exclude partial first/last observation months, include intervening zero months.
                observed = bool(dates[client] and min(dates[client]) <= start and max(dates[client]) >= end)
                status, activity = ('full', 'observed_span_estimate') if observed else ('unknown', 'unknown')
                allocation, basis = weights.get(client), 'estimated'
            eligible = status == 'full' and allocation is not None
            if not rows and not entry and status != 'full':
                continue
            counts = [sum(r['band'] == i for r in rows) for i in range(8)]
            included = [counts[i] if eligible and allocation[i] > 0 else 0 for i in range(8)]
            audit.append({'client': client, 'month': period, 'activity': activity, 'allocation_source': basis,
                          'eligible': eligible, 'reason': 'included' if eligible else 'partial/inactive/unknown activity or allocation',
                          'weights': allocation, 'band_leads': included, 'recorded_band_leads': counts,
                          'unallocated_leads': sum(counts)-sum(included),
                          'unknown_size_leads': sum(r['band'] is None for r in rows),
                          'evidence': entry['evidence'] if entry else 'Undated backend weights; bounded first/last qualifying lead observations (activity proxy, not a verified campaign schedule).'})

    def summary(indices, months):
        eligible = [a for a in audit if a['month'] in months and a['eligible'] and sum(a['weights'][i] for i in indices) > 0]
        totals, per_client = Counter(), defaultdict(lambda: [0, 0.0])
        for a in eligible:
            exposure = sum(a['weights'][i] for i in indices)
            leads = sum(a['band_leads'][i] for i in indices)
            totals['leads'] += leads
            totals['exposure'] += exposure
            totals[a['allocation_source']] += exposure
            totals['verified_months'] += a['activity'] == 'verified'
            per_client[a['client']][0] += leads
            per_client[a['client']][1] += exposure
        exposure, leads = totals['exposure'], totals['leads']
        rate = leads/exposure if exposure else None
        sufficient = len(per_client) >= MIN_CLIENTS and exposure >= MIN_EXPOSURE
        without = [(leads-v[0])/(exposure-v[1]) for v in per_client.values() if exposure-v[1] > 1e-9]
        estimated_activity = sum(a['activity'] != 'verified' for a in eligible)
        return {'average': rate, 'clients': len(per_client), 'qualifying_leads': leads,
                'equivalent_months': exposure, 'verified_live_client_months': totals['verified_months'],
                'estimated_live_client_months': estimated_activity, 'sufficient': sufficient,
                'allocation_percent': {k: 100*totals[k]/exposure if exposure else 0 for k in ('outreach','approved','estimated')},
                'dominant_client_lead_share': max((v[0]/leads for v in per_client.values()), default=0) if leads else 0,
                'leave_one_client_out': [min(without), max(without)] if without else [],
                'estimated': bool(estimated_activity or totals['estimated']), 'months': months}
    groups = []
    for label, headcount, indices in SEGMENTS:
        result = summary(indices, base)
        subs = [summary((i,), result['months']) for i in indices]
        result.update(label=label, headcount=headcount, bands=list(indices))
        display = 'Insufficient data'
        if result['sufficient']:
            value = result['average']
            display = f'~{math.floor(value+0.5)}' if value >= 0.5 else ('<1' if value > 0 else '0')
            if len(subs) > 1 and all(s['sufficient'] for s in subs):
                rates = [s['average'] for s in subs]
                low, high = math.floor(min(rates)+0.5), math.floor(max(rates)+0.5)
                if high > low and max(rates)-min(rates) >= 1:
                    display = f'{low}–{high}'
        result['display'] = display
        result['confidence'] = 'Insufficient sample' if not result['sufficient'] else 'Estimated inputs' if result['estimated'] else 'Supported sample'
        if result['dominant_client_lead_share'] > .5:
            result['confidence'] += ' · one client contributes over half the leads'
        groups.append(result)
    band_results = []
    for i, prior in enumerate((11.0,8.1,5.0,3.4,2.8)):
        r = summary((i,), base)
        band_results.append({'label': BANDS[i], 'preliminary': prior, **r,
                             'difference': r['average']-prior if r['average'] is not None else None})
    ent = summary((5, 6, 7), base)
    band_results.append({'label': '1,001+', 'preliminary': 3.0, **{k:v for k,v in ent.items() if k not in ('label','display')},
                         'difference': ent['average']-3 if ent['average'] is not None else None})
    quality.update({'unmatched_allocation_or_activity_client_months': sum(not a['eligible'] for a in audit),
                    'unallocated_segment_leads': sum(a['unallocated_leads'] for a in audit),
                    'reconciliation_mismatches': sum(r['difference'] != 0 for r in reconciliation),
                    'reconciliation_compared_client_months': len(reconciliation)})
    return {'groups': groups, 'bands': band_results, 'quality': quality, 'warnings': warnings,
            'period': base, 'audit': audit, 'reconciliation': reconciliation,
            'methodology_version': 3, 'minimum_clients': MIN_CLIENTS, 'minimum_equivalent_months': MIN_EXPOSURE,
            'reference_estimates_note': 'Differences reflect rolling periods, exclusions, inclusion of Interested and zero-lead exposure, normalised targeting allocations and small-sample thresholds. Estimates are not fitted to preliminary figures.'}
