"""Validation for the optional client growth planner. No reporting data is mutated."""
from datetime import date
import math
import uuid

MAX_AMOUNT = 1_000_000_000


def number(value, label, maximum=MAX_AMOUNT):
    if isinstance(value, bool):
        raise ValueError(f'{label} must be a number.')
    try:
        value = float(value)
    except (ValueError, TypeError):
        raise ValueError(f'Enter {label.lower()}.') from None
    if not math.isfinite(value) or value < 0 or value > maximum:
        raise ValueError(f'{label} must be between 0 and {maximum:,.0f}.')
    return value


def day(value, label):
    try:
        return date.fromisoformat(str(value))
    except (TypeError, ValueError):
        raise ValueError(f'Enter a valid {label.lower()}.') from None


def validate(data, today):
    if not isinstance(data, dict):
        raise ValueError('Enter your business figures.')
    mode = data.get('mode', 'monthly')
    if mode not in ('monthly', 'project'):
        raise ValueError('Choose Monthly or Project.')
    currency = data.get('currency', 'GBP')
    if currency not in ('GBP', 'USD', 'EUR'):
        raise ValueError('Choose GBP, USD or EUR.')
    if data.get('target_basis', 'leads') not in ('leads', 'sales'):
        raise ValueError('Choose a sales or lead target.')
    start = day(data.get('measurement_start'), 'lead tracking start date')
    if start > today:
        raise ValueError('Lead tracking cannot start in the future.')
    conversions = data.get('conversions', [])
    if not isinstance(conversions, list) or len(conversions) > 500:
        raise ValueError('Enter at most 500 conversions.')
    clean = []
    ids = set()
    for row in conversions:
        if not isinstance(row, dict):
            raise ValueError('Check each conversion.')
        landed = day(row.get('landed'), 'conversion date')
        if landed < start or landed > today:
            raise ValueError('Conversion dates must fall between the tracking start date and today.')
        end_month = row.get('end_month') or ''
        if end_month:
            end = day(str(end_month)+'-01', 'retainer end month')
            if str(end_month) != end.strftime('%Y-%m') or end.strftime('%Y-%m') < landed.strftime('%Y-%m') or end.strftime('%Y-%m') > today.strftime('%Y-%m'):
                raise ValueError('End month must be between the landed month and this month.')
        amount = number(row.get('amount'), 'Conversion amount')
        if amount <= 0:
            raise ValueError('Each conversion needs an amount above zero.')
        key = str(row.get('id') or uuid.uuid4())
        if len(key) > 80 or key in ids:
            raise ValueError('Each conversion must have a unique ID.')
        ids.add(key)
        clean.append({'id': key, 'landed': landed.isoformat(), 'amount': amount, 'end_month': end_month,
                      'name': str(row.get('name') or '')[:120]})
    extra = {}
    if data.get('schema_version') == 2:
        goals = data.get('goals')
        if not isinstance(goals, list) or len(goals) != 5:
            raise ValueError('Enter up to five yearly revenue goals.')
        extra = {'schema_version': 2,
                 'average_value': number(data.get('average_value'), 'Average customer value'),
                 'goals': [None if v is None or v == '' else number(v, 'Yearly revenue goal') for v in goals]}
    return {
        **extra,
        'mode': mode, 'currency': currency, 'measurement_start': start.isoformat(),
        'annual_churn': number(data.get('annual_churn'), 'Annual churn', 100),
        'current_revenue': number(data.get('current_revenue'), 'Current revenue'),
        'goal_12': number(data.get('goal_12'), '12-month revenue goal'),
        'goal_24': number(data.get('goal_24'), '24-month revenue goal'),
        'annual_sales': number(data.get('annual_sales', 0), 'Annual sales target', 12_000_000),
        'target_basis': data.get('target_basis', 'leads'),
        'monthly_leads': number(data.get('monthly_leads'), 'Monthly lead target', 1_000_000),
        'conversions': clean,
    }


def defaults(today):
    return {'mode': 'monthly', 'currency': 'GBP', 'measurement_start': today.replace(day=1).isoformat(),
            'annual_churn': 0, 'current_revenue': 0, 'goal_12': 0, 'goal_24': 0,
            'monthly_leads': 0, 'annual_sales': 0, 'target_basis': 'leads', 'conversions': []}
