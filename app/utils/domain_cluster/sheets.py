"""Fresh read-only inventory from the two explicitly authorised strategy tabs.

No content cache: every tool submission performs a new Google API read. Only A:G
is requested, deliberately excluding recovery addresses and codes in later columns.
Mailbox local parts are discarded before any response, logging or persistence.
"""
from __future__ import annotations

import copy

import requests
from fastapi import HTTPException
from google.auth.exceptions import GoogleAuthError
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials

from app.utils import google_sheets as gs
from . import engine as e

SHEET_ID = '1PN5Lx-5c4nbouDl-Apw9YJgrYTuoDsBKqzk5Hh6Z0po'
SOURCE_URL = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit'
TABS = {'client': (0, '1) Emails - entry'), 'bizdev': (838806676, '1) BD Emails - entry')}
MAX_ROWS = 10000


def _session():
    try:
        # Reuse the established credential binding with read-only scopes. Never
        # alter the shared Google client or print the decoded credential.
        credentials = Credentials.from_service_account_info(gs._decode_sa_json(), scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
        return AuthorizedSession(credentials)
    except Exception:
        raise HTTPException(503, 'Strategy sheet access is not configured. Ask an administrator to check the existing Google Sheets service-account connection.') from None


def checked(value):
    # bool('FALSE') is True: never use generic truthiness for checkboxes.
    return value is True or isinstance(value, str) and value.strip().upper() == 'TRUE'


def parse_tab(kind, title, values):
    if not values:
        raise HTTPException(502, f'The {title} tab has no header. No inventory changes were applied.')
    header = next(((i, [str(v).strip().lower() for v in row]) for i, row in enumerate(values[:10]) if any(str(v).strip().lower() == 'live' for v in row) and any(str(v).strip().lower().startswith('email') for v in row)), None)
    if not header:
        raise HTTPException(502, f'Email and Live columns were not found in {title}. No inventory changes were applied.')
    offset, names = header
    email_col = next(i for i,h in enumerate(names) if h.startswith('email'))
    live_col = names.index('live')
    owner_col = names.index('biz dev lead') if 'biz dev lead' in names else names.index('client') if 'client' in names else None
    def cell(row, col):
        return row[col] if col is not None and col < len(row) else ''
    records, invalid, scanned, live = [], [], 0, 0
    for row_number, row in enumerate(values[offset+1:], offset+2):
        if not any(str(v).strip() for v in row):
            continue
        scanned += 1
        if not checked(cell(row, live_col)):
            continue
        live += 1
        try:
            domain = e.normalize(str(cell(row, email_col)))
        except ValueError:
            invalid.append({'sheet': title, 'row': row_number, 'reason': 'Live row has an invalid or missing email/domain. Correct it in the strategy sheet.'})
            continue
        flags = {key: checked(cell(row, names.index(label))) if label in names else False for key, label in [('sheet_at_risk','at risk'), ('sheet_needs_deleting','needs deleting'), ('sheet_replacing','being replaced')]}
        owner = str(cell(row, owner_col)).strip()[:100]
        if '@' in owner:
            owner = ''  # Do not persist unexpected personal addresses from metadata.
        records.append({'domain':domain,'source_kind':kind,'source_tab':title,'source_row':row_number,'source_owner':owner,**flags})
    return {'records':records,'invalid':invalid,'summary':{'kind':kind,'tab':title,'scanned_rows':scanned,'live_rows':live,'valid_live_rows':len(records),'invalid_live_rows':len(invalid),'unique_domains':len({r['domain'] for r in records}),'at_risk_rows':sum(r['sheet_at_risk'] for r in records),'needs_deleting_rows':sum(r['sheet_needs_deleting'] for r in records)}}


def read_live():
    """Metadata + one bounded batchGet per invocation; a partial read fails closed."""
    session = _session()
    base = f'https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}'
    try:
        def get(path, params):
            # Share the app's Sheets read accounting/pacing, but do not retry or
            # cache snapshots: each submission must either get fresh data or fail.
            gs._read_limiter.acquire()
            gs._bump_sheet_reads()
            response = session.get(path, params=params, timeout=(5,20), headers={'Cache-Control':'no-cache'})
            if response.status_code in {401,403,404}:
                raise HTTPException(503, 'The strategy sheet could not be opened. Confirm it is shared as Viewer with the dashboard Google Sheets service account.')
            if response.status_code == 429:
                raise HTTPException(429, 'Google Sheets read quota reached. Retry shortly; no cached inventory was used.')
            response.raise_for_status()
            return response.json()
        metadata = get(base, {'fields':'sheets.properties'})
        tabs = [r['properties'] for r in metadata.get('sheets',[])]
        resolved, ranges = [], []
        for kind, (gid, title) in TABS.items():
            tab = next((t for t in tabs if t.get('sheetId') == gid), None)
            if tab is None or tab.get('title') != title:
                raise HTTPException(502, f'The source tab {title} is missing or renamed. No inventory changes were applied.')
            row_count = tab.get('gridProperties',{}).get('rowCount',0)
            if not 1 <= row_count <= MAX_ROWS:
                raise HTTPException(502, 'The strategy sheet exceeds the supported row limit. No partial inventory was imported.')
            resolved.append((kind,title))
            ranges.append("'"+title.replace("'","''")+f"'!A1:G{row_count}")
        data = get(base+'/values:batchGet', [('ranges',r) for r in ranges]+[('valueRenderOption','UNFORMATTED_VALUE')])
        values = data.get('valueRanges',[])
        if len(values) != len(resolved):
            raise HTTPException(502, 'Google returned an incomplete inventory. Retry; no cached or partial inventory was used.')
        results = [parse_tab(kind,title,raw.get('values',[])) for (kind,title),raw in zip(resolved,values)]
        # An unreadable Live row can hide a capacity-consuming domain. Do not
        # approve recommendations against an incomplete current inventory.
        invalid = [row for tab in results for row in tab['invalid']]
        return {'records':[row for tab in results for row in tab['records']], 'invalid':invalid, 'tabs':[tab['summary'] for tab in results], 'checked_at':e.now(), 'source_url':SOURCE_URL}
    except (requests.RequestException, GoogleAuthError, KeyError, ValueError):
        raise HTTPException(503, 'Fresh strategy sheet scan failed. Retry; the previous inventory was not used for this request.') from None
    finally:
        session.close()


def refresh(state, source):
    """Build a reviewable working inventory; never writes the sheet or portfolio."""
    state = copy.deepcopy(state)
    by_domain = {}
    for r in source['records']:
        by_domain.setdefault(r['domain'], []).append(r)
    if not by_domain:
        raise HTTPException(409, 'No valid Live domains were found. Check the sheet before replacing any inventory.')
    existing = {r['domain']:r for r in state['domains']}
    incoming = set(by_domain)
    missing = sorted(r['domain'] for r in state['domains'] if r.get('sheet_sources') and r['status']=='Active' and r['domain'] not in incoming)
    added = sorted(incoming - set(existing))
    changed = []
    for domain, rows in by_domain.items():
        if domain not in existing:
            row={'domain':domain,'status':'Active','owner':'','platform':'','notes':'','added_at':e.now(),'override':None,'reviewed':False}
            state['domains'].append(row)
        else:
            row = existing[domain]
            if row['status'] != 'Active':
                changed.append(domain)
            row['status']='Active'
        row.update(sheet_sources=sorted({r['source_kind'] for r in rows}),sheet_owners=sorted({r['source_owner'] for r in rows if r['source_owner']}),sheet_at_risk=any(r['sheet_at_risk'] for r in rows),sheet_needs_deleting=any(r['sheet_needs_deleting'] for r in rows),sheet_replacing=any(r['sheet_replacing'] for r in rows),source_checked_at=source['checked_at'])
    for domain in missing:
        existing[domain]['status']='Paused'
    if len(state['domains']) > e.MAX_DOMAINS:
        raise HTTPException(400, 'The combined portfolio exceeds 5,000 domains. No partial inventory was imported.')
    preview={'added':added,'existing':sorted(incoming & set(existing)),'missing':missing,'reactivated':changed,'mode':'sheet','invalid_count':len(source['invalid']),'duplicates':len(source['records'])-len(incoming)}
    return state, preview
