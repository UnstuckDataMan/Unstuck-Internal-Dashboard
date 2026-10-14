"""Transparent peak account-month examples, separate from cohort averages."""
from collections import Counter


def build_examples(records, months, groups, allowed=None):
    eligible = [r for r in records if r['date'].strftime('%Y-%m') in months
                and (allowed is None or r['client'] in allowed)]
    clients = sorted({r['client'] for r in eligible if r['client'] != 'unstuck'})
    references = {client: f'Client example {i+1:02d}' for i,client in enumerate(clients)}
    result = []
    for index, (label, bands) in enumerate(groups):
        counts = Counter()
        for row in eligible:
            if row['band'] in bands:
                counts[row['client'], row['date'].strftime('%Y-%m')] += row.get('leads', 1)
        # Highest recorded count, then most recent month, then stable account order.
        positive = [key for key in counts if counts[key] > 0]
        if not positive:
            result.append({'label': label, 'key': f'group-{index}', 'leads': None,
                           'candidate_accounts': 0, 'observed_account_months': 0})
            continue
        account, month = sorted(positive, key=lambda k: (-counts[k], -int(k[1].replace('-','')), k[0]))[0]
        result.append({'label': label, 'key': f'group-{index}', 'leads': counts[account,month],
                       'month': month, 'account_label': 'Unstuck — own prospecting' if account == 'unstuck' else references[account],
                       'is_unstuck': account == 'unstuck', 'sample_accounts': 1,
                       'candidate_accounts': len({a for a,m in positive}),
                       'observed_account_months': len(positive)})
    return {'example_window': months, 'examples': result}
