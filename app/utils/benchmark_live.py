"""Half-hourly, server-side benchmark snapshots; isolated read-only Sheets session."""
from copy import deepcopy
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import tempfile
import threading
import time
from urllib.parse import quote

import requests

from app.utils import client_benchmarks as source
from app.utils import benchmark_model as model
from app.utils.google_sheets import _read_limiter

INTERVAL = 1800
TABS = {'leads': (1768027115, 'Form Responses'), 'monthly': (830915082, 'Monthly leads'),
        'weights': (786537287, 'Lead Data Averages Backend'), 'previous': (65292198, 'Lead Data Averages')}
ERI_SHEET = '1UIb0HF1LlnD4lR5nV9RG9hg_fX8PHHdPCRIgZVAU1lQ'
_lock = threading.Lock()
_snapshot = None
_attempt = 0.0
_error = None
_stop = threading.Event()
_thread = None


def now():
    return datetime.now(timezone.utc).isoformat()


def cache_path():
    return Path(os.environ.get('CLIENT_BENCHMARK_CACHE_PATH', f'/tmp/unstuck-client-benchmarks-v{model.METHODOLOGY_VERSION}.json'))


def read_response(session, url, params):
    _read_limiter.acquire()
    r = session.get(url, params=params, timeout=(5, 25))
    r.raise_for_status()
    return r.json()


def fetch_tabs():
    if not source._is_configured():
        raise source.SourceError('Google Sheets credentials are not configured for this server.')
    session = source._read_session()
    try:
        base = f'https://sheets.googleapis.com/v4/spreadsheets/{source.SHEET_ID}'
        metadata = read_response(session, base, {'fields': 'sheets.properties'})
        sheets = [s['properties'] for s in metadata.get('sheets', [])]
        selected, notices = {}, []
        for key, (gid, name) in TABS.items():
            tab = next((s for s in sheets if s['sheetId'] == gid), None)
            if not tab:
                tab = next((s for s in sheets if s['title'].strip().lower() == name.lower()), None)
                if tab:
                    notices.append(f'{name}: matched by title because the supplied tab ID was not found.')
            if tab:
                selected[key] = tab
            elif key in ('leads', 'monthly', 'weights'):
                raise source.SourceError(f'Required source tab missing: {name}.')
        control = next((s for s in sheets if s['title'].strip().lower() == 'benchmark campaign months'), None)
        if control:
            selected['controls'] = control
        else:
            notices.append('No Benchmark campaign months tab: activity and backend allocations are estimates where available.')
        ranges = ["'"+s['title'].replace("'", "''")+"'!A:AZ" for s in selected.values()]
        response = read_response(session, base+'/values:batchGet', {'ranges': ranges, 'valueRenderOption': 'UNFORMATTED_VALUE', 'dateTimeRenderOption': 'SERIAL_NUMBER'})
        values = response.get('valueRanges', [])
        if len(values) != len(selected):
            raise source.SourceError('Google returned an incomplete source snapshot. The previous successful results are retained.')
        tabs = {key: item.get('values', []) for key, item in zip(selected, values)}
        # Reference only. Never append ERI tracker data to benchmark lead records.
        try:
            eri_base = f'https://sheets.googleapis.com/v4/spreadsheets/{ERI_SHEET}'
            eri_meta = read_response(session, eri_base, {'fields': 'sheets.properties'})
            eri_tabs = [s['properties'] for s in eri_meta.get('sheets', [])]
            reference = next((s for s in eri_tabs if s['title'].lower() == 'form responses'), None)
            if reference:
                eri = read_response(session, eri_base+'/values/'+quote("'"+reference['title'].replace("'", "''")+"'!A:AZ", safe=''), {'valueRenderOption': 'UNFORMATTED_VALUE'})
                # ERI is excluded from all model inputs, including this reference.
                tabs['eri_reference'] = eri.get('values', [])
                eri_note = 'ERI reference retrieved; excluded from benchmark inputs. Tracker schema requires a separate reviewed reconciliation.'
            else:
                eri_note = 'ERI tracker accessible; no comparable Form Responses tab. Automatic ERI reconciliation is unavailable. ERI remains excluded.'
        except Exception:
            eri_note = 'ERI reference could not be accessed; excluded from all benchmark calculations. Reference cross-check remains unverified.'
        return tabs, {'tabs': {k: {'title': v['title'], 'gid': v['sheetId']} for k,v in selected.items()},
                      'notices': notices, 'eri_reference': eri_note}
    finally:
        session.close()


def safe_error(exc):
    if isinstance(exc, source.SourceError):
        return str(exc)
    if isinstance(exc, requests.HTTPError):
        status = exc.response.status_code if exc.response is not None else None
        if status in (401,403,404):
            return 'Google denied access. Check the existing service account has Viewer access to Lead Database and Sheets API is enabled.'
        if status == 429:
            return 'Google Sheets rate limit reached. The next scheduled sync will retry.'
    return 'Source synchronisation failed. Check Google connectivity, credentials and the campaign-month control tab.'


def restore():
    global _snapshot
    if _snapshot is not None:
        return
    try:
        loaded = json.loads(cache_path().read_text())
        if loaded.get('methodology_version') == model.METHODOLOGY_VERSION and loaded.get('synced_at') and loaded.get('groups'):
            _snapshot = loaded
    except (OSError, ValueError, TypeError):
        pass


def persist(snapshot):
    path = cache_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    # Snapshot includes internal audit names: restrict to the service user.
    name = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', dir=path.parent, delete=False) as f:
            name = f.name
            json.dump(snapshot, f, allow_nan=False)
        os.replace(name, path)
    finally:
        if name and os.path.exists(name):
            os.unlink(name)


def refresh():
    global _snapshot, _attempt, _error
    with _lock:
        restore()
        if _attempt and time.monotonic()-_attempt < INTERVAL:
            return
        _attempt = time.monotonic()
        try:
            tabs, metadata = fetch_tabs()
            synced = now()
            result = model.calculate(tabs)
            result.update(synced_at=synced, calculated_at=now(), source_metadata=metadata)
            try:
                persist(result)
            except OSError:
                result['warnings'].append('Disk snapshot unavailable; last successful data is retained in this process only.')
            _snapshot, _error = result, None
        except Exception as exc:
            _error = safe_error(exc)


def get_report(internal=False):
    refresh()
    with _lock:
        if _snapshot is None:
            raise source.SourceError(_error or 'No successful source synchronisation is available yet.')
        result = deepcopy(_snapshot)
        # Reformat compatible cached calculations without discarding a last-good snapshot.
        for group in result['groups']:
            group['display'] = model.display_rate(group)
        age = (datetime.now(timezone.utc)-datetime.fromisoformat(result['synced_at'])).total_seconds()
        result.update(stale=bool(_error or age > INTERVAL*2), sync_error=_error, polling_minutes=30)
        if age > 7*86400:
            for group in result['groups']:
                group['display'] = 'Data too old'
            result['stale'] = True
        if not internal:
            result.pop('audit', None)
            result.pop('reconciliation', None)
        return result


def start_scheduler():
    global _thread
    if not source._is_configured() or (_thread and _thread.is_alive()):
        return
    _stop.clear()
    def run():
        if _stop.wait(10):
            return
        while not _stop.is_set():
            refresh()
            if _stop.wait(INTERVAL):
                break
    _thread = threading.Thread(target=run, name='client-benchmark-sync', daemon=True)
    _thread.start()


def stop_scheduler():
    _stop.set()
