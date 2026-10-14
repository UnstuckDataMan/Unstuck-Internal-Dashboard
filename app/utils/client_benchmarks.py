"""Read-only lead benchmarks. Never infer campaign activity from a lead date."""
from __future__ import annotations

import base64
import json
import os
import csv
import io
import re
import threading
import time
from collections import Counter
from datetime import date, datetime, timezone
from statistics import median
from urllib.parse import quote
from zoneinfo import ZoneInfo

import requests

from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials

SHEET_ID = "1mhJ0lytLPmfkhMstOy2aF6FMjZPoWalxIigfK18ZqsY"
GID = 2084964985
SOURCE_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit#gid={GID}"
BANDS = ("1–10", "11–50", "51–200", "201–500", "501–1,000", "1,001–5,000", "5,001–10,000", "10,001+")
LIMITS = (10, 50, 200, 500, 1000, 5000, 10000)
GROUPS = (("Micro", (0,)), ("SME", (1, 2)), ("Medium", (3, 4)), ("Enterprise", (5, 6, 7)))
HEADERS = ("client", "date", "lead", "job role", "company size", "company industry", "channel", "country")
_cache = None
_lock = threading.Lock()


class SourceError(Exception):
    pass


def london_today():
    return datetime.now(ZoneInfo("Europe/London")).date()


def default_period(today=None):
    today = today or london_today()
    end_index = today.year * 12 + today.month - 2
    return month_string(end_index - 2), month_string(end_index)


def month_string(index):
    year, month = divmod(index, 12)
    return f"{year:04d}-{month + 1:02d}"


def period_months(start, end, today=None):
    today = today or london_today()
    if not all(re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", s) for s in (start, end)):
        raise ValueError("Choose valid start and end months.")
    def index(s):
        y, m = map(int, s.split("-"))
        return y * 12 + m - 1
    first, last = index(start), index(end)
    if first > last or last - first >= 36:
        raise ValueError("Choose a period of 1 to 36 months, with the start before the end.")
    if last >= today.year * 12 + today.month - 1:
        raise ValueError("Use completed months so partial-month results do not distort the average.")
    return [month_string(i) for i in range(first, last + 1)]


def _normal(value):
    return " ".join(str(value or "").strip().lower().split())


def parse_date(value):
    value = str(value or "").strip()
    # ISO first; ambiguous slash dates follow the dashboard's UK convention.
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        if parsed.tzinfo:
            parsed = parsed.astimezone(ZoneInfo("Europe/London"))
        return parsed.date()
    except ValueError:
        pass
    for fmt in ("%d/%m/%Y", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%y", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def band_index(value):
    value = re.sub(r"\s+", "", str(value or "")).replace(",", "").replace("–", "-").replace("—", "-")
    canonical = [re.sub(r"[\s,]", "", b).replace("–", "-") for b in BANDS]
    if value in canonical:
        return canonical.index(value)
    if value.isdigit() and int(value) > 0:
        count = int(value)
        return next((i for i, upper in enumerate(LIMITS) if count <= upper), 7)
    # Do not split an ambiguous range, e.g. 50–200, across two bands.
    return None


def parse_rows(values):
    for header_index, row in enumerate(values[:10]):
        names = [_normal(cell).lstrip("\ufeff") for cell in row]
        if all(name in names for name in HEADERS):
            if any(names.count(name) != 1 for name in HEADERS):
                raise SourceError("The source contains duplicate column headers. Check the linked sheet.")
            columns = [names.index(name) for name in HEADERS]
            break
    else:
        raise SourceError("The linked tab must contain Client, Date, Lead, Job role, Company size, Company Industry, Channel and Country columns.")
    records, seen = [], set()
    quality = {"invalid_rows": 0, "duplicates": 0, "unknown_size": 0}
    for row in values[header_index + 1:]:
        cells = tuple(str(row[i]).strip() if i < len(row) else "" for i in columns)
        if not any(cells):
            continue
        client, raw_date, lead, role, size, industry, channel, country = cells
        day = parse_date(raw_date)
        if not client or not day or not lead:
            quality["invalid_rows"] += 1
            continue
        # Identical source rows only. Different people/roles/dates remain distinct.
        signature = tuple(_normal(v) for v in cells)
        if signature in seen:
            quality["duplicates"] += 1
            continue
        seen.add(signature)
        band = band_index(size)
        if band is None:
            quality["unknown_size"] += 1
        records.append({"client": _normal(client), "date": day, "band": band})
    return records, quality


def _is_configured():
    return bool(os.environ.get('GOOGLE_SHEETS_SA_JSON', '').strip())


def _read_session():
    """Independent read-only session; never mutate other tools' auth caches."""
    raw = os.environ.get('GOOGLE_SHEETS_SA_JSON', '').strip()
    info = json.loads(base64.b64decode(raw + '=='))
    credentials = Credentials.from_service_account_info(
        info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
    return AuthorizedSession(credentials)


def load_source():
    """Server-side access only. Reuse the dashboard's existing Sheets binding."""
    global _cache
    with _lock:
        if _cache is not None and time.monotonic() - _cache[0] < 60:
            return _cache[1]
        try:
            if _is_configured():
                session = _read_session()
                base = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}"
                metadata = session.get(base, params={"fields": "sheets.properties"}, timeout=20)
                metadata.raise_for_status()
                title = next((s["properties"]["title"] for s in metadata.json().get("sheets", []) if s["properties"]["sheetId"] == GID), None)
                if title is None:
                    raise SourceError("The selected source tab could not be found.")
                sheet_range = "'" + title.replace("'", "''") + "'!A:Z"
                response = session.get(base + "/values/" + quote(sheet_range, safe=""), params={"valueRenderOption": "FORMATTED_VALUE"}, timeout=20)
                response.raise_for_status()
                values = response.json().get("values", [])
            else:
                response = requests.get(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export", params={"format": "csv", "gid": GID}, timeout=20)
                response.raise_for_status()
                if "text/html" in response.headers.get("Content-Type", ""):
                    raise SourceError("The source sheet requires access. Connect the dashboard's Google Sheets service account and share this sheet with it as a viewer.")
                values = list(csv.reader(io.StringIO(response.content.decode("utf-8-sig"))))
            records, quality = parse_rows(values)
        except SourceError:
            raise
        except Exception as exc:
            # Never expose OAuth exceptions, response bodies or credentials to clients.
            raise SourceError("The live sheet is unavailable. Check Google network access and share the sheet with the dashboard's configured Google Sheets service account. No example figures are being substituted.") from exc
        result = {"records": records, "quality": quality, "fetched_at": datetime.now(timezone.utc).isoformat()}
        _cache = (time.monotonic(), result)
        return result


def summarize(source, months):
    if 'delivery_roster' in source:
        from app.utils.benchmark_billing import summarize as summarize_billing
        return summarize_billing(source)
    active = 'active_records' in source
    selected = [r for r in source.get('active_records', source.get('records', [])) if r['date'].strftime('%Y-%m') in months]
    def segment(key, label, indices):
        rows = [r for r in selected if r['band'] in indices]
        clients = {r['client'] for r in rows}
        subject = lambda r: r['campaign'] if active else r['client']
        counts = Counter()
        monthly = Counter()
        exposure = set()
        for r in rows:
            m = r['date'].strftime('%Y-%m')
            n = r.get('leads', 1)
            counts[subject(r)] += n
            monthly[m] += n
            exposure.add((subject(r), m))
        if not active:
            exposure = {(client, m) for client in clients for m in months}
        subject_months = Counter(s for s, m in exposure)
        month_counts = Counter(m for s, m in exposure)
        denominator = len(exposure)
        total = sum(counts.values())
        rates = [monthly[m] / month_counts[m] for m in months if month_counts[m]]
        return {
            'key': key, 'label': label, 'bands': list(indices), 'leads': total,
            'clients': len(clients), 'campaigns': len(counts) if active else None,
            'client_months': denominator, 'exposure_months': denominator,
            'unstuck_leads': sum(r.get('leads', 1) for r in rows if r['client'] == 'unstuck'),
            'average': total / denominator if denominator else None,
            'median': median([n / subject_months[s] for s, n in counts.items()]) if counts else None,
            'best_month': max(rates) if rates else None,
            'monthly': [{'month': m, 'leads': monthly[m], 'exposure_months': month_counts[m],
                         'average': monthly[m] / month_counts[m] if month_counts[m] else None} for m in months],
        }
    from app.utils.benchmark_examples import build_examples
    return {
        **build_examples(selected, months, GROUPS),
        'months': months, 'fetched_at': source['fetched_at'], 'source_url': SOURCE_URL,
        'basis': 'active_campaign_month' if active else 'calendar_account_month',
        **{key: source[key] for key in ('source_kind', 'source_tab', 'date_min', 'date_max', 'lead_definition') if key in source},
        'bands': [segment(f'band-{i}', label, (i,)) for i, label in enumerate(BANDS)],
        'groups': [segment(f'group-{i}', label, indices) for i, (label, indices) in enumerate(GROUPS)],
        'overall': segment('all', 'All target sizes', tuple(range(8))),
        'quality': source['quality'],
        'unmapped_in_period': sum(r.get('leads', 1) for r in selected if r['band'] is None),
    }
