"""Ratios to supplied lifetime delivery-month totals, never date-window averages."""
from collections import Counter
from statistics import median

from app.utils import client_benchmarks as b

HEADERS = ('client', 'lead sheet client', 'delivery months')


def read_roster(workbook):
    if 'Delivery months' not in workbook.sheetnames:
        return None
    rows = workbook['Delivery months'].iter_rows(max_col=3, values_only=True)
    names = [b._normal(x) for x in next(rows, ())]
    if tuple(names) != HEADERS:
        raise b.SourceError('Delivery months headers must be Client, Lead sheet client, Delivery months.')
    roster, seen = [], set()
    for row in rows:
        if not any(x is not None for x in row):
            continue
        client, alias, count = row
        client = b._normal(client)
        alias = b._normal(alias) or client
        if not client or alias in seen or isinstance(count, bool):
            raise b.SourceError('Delivery-month rows require unique client mappings.')
        try:
            number = float(count)
            if not number.is_integer() or number <= 0 or number > 10000:
                raise ValueError()
        except (TypeError, ValueError):
            raise b.SourceError('Delivery months must be a positive whole number.')
        seen.add(alias)
        roster.append({'client': client, 'alias': alias, 'months': int(number)})
        if len(roster) > 10000:
            raise b.SourceError('The delivery-month roster is too large.')
    if not roster:
        raise b.SourceError('The Delivery months tab is empty.')
    return roster


def summarize(source):
    roster = {r['alias']: r['months'] for r in source['delivery_roster']}
    records = [r for r in source['records'] if r['client'] in roster]
    matched = {r['client'] for r in records}
    if not matched:
        raise b.SourceError('No delivery-month clients match the lead records. Check the Lead sheet client mappings.')
    def segment(key, label, bands):
        rows = [r for r in records if r['band'] in bands]
        # Only accounts observed in a segment can be assigned to it. No activity inferred.
        counts = Counter(r['client'] for r in rows)
        denominator = sum(roster[c] for c in counts)
        return {'key': key, 'label': label, 'bands': list(bands), 'leads': len(rows),
                'clients': len(counts), 'campaigns': None,
                'client_months': denominator, 'exposure_months': denominator,
                'unstuck_leads': counts.get('unstuck', 0),
                'average': len(rows)/denominator if denominator else None,
                'median': median([n/roster[c] for c,n in counts.items()]) if counts else None,
                'best_month': None, 'monthly': []}
    from app.utils.benchmark_examples import build_examples
    end_year, end_month = map(int, b.default_period()[1].split('-'))
    end_index = end_year * 12 + end_month - 1
    example_months = [b.month_string(i) for i in range(end_index-11, end_index+1)]
    examples = build_examples(source['records'], example_months, b.GROUPS, set(roster) | {'unstuck'})
    metadata = {k: source[k] for k in ('fetched_at','source_kind','source_tab','date_min','date_max','lead_definition') if k in source}
    return {**metadata, **examples, 'source_url': b.SOURCE_URL, 'basis': 'delivery_month',
            'months': [min(r['date'] for r in records).strftime('%Y-%m'),max(r['date'] for r in records).strftime('%Y-%m')],
            'roster_accounts': len(roster), 'matched_accounts': len(matched),
            'unmatched_accounts': len(roster)-len(matched),
            'unmatched_delivery_months': sum(n for c,n in roster.items() if c not in matched),
            'unstuck_in_roster': 'unstuck' in roster,
            'bands': [segment(f'band-{i}',label,(i,)) for i,label in enumerate(b.BANDS)],
            'groups': [segment(f'group-{i}',label,indices) for i,(label,indices) in enumerate(b.GROUPS)],
            'overall': segment('all','All target sizes',tuple(range(8))),
            'quality': source['quality'], 'unmapped_in_period': sum(r['band'] is None for r in records)}
