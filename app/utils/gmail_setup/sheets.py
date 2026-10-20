"""Idempotent completion publishing to the two authorised strategy tabs only.

Read headings, email identifiers and operational columns; never recovery codes.
Create a new native row atomically with its marker, typed values and checkboxes.
A retry locates the stable row marker or exact email before attempting an insert.
"""
import hashlib
import requests
from fastapi import HTTPException
from app.utils import google_sheets as gs
from app.utils.domain_cluster.sheets import SHEET_ID, SOURCE_URL, TABS, checked
from app.utils.strategy_sheet_identity import email_cell_identity

BASE = f'https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}'


def col(index):
    result = ''
    index += 1
    while index:
        index, rem = divmod(index-1, 26)
        result = chr(65+rem)+result
    return result


def a1(title, cells):
    return "'"+title.replace("'", "''")+"'!"+cells


def call(session, method, url, guard=None, **kwargs):
    if method == 'get':
        gs._read_limiter.acquire()
        gs._bump_sheet_reads()
    else:
        gs._write_limiter.acquire()
    if guard is not None:
        guard()
    try:
        response = getattr(session, method)(url, timeout=(3, 15), **kwargs)
        if response.status_code in {401, 403}:
            raise HTTPException(503, 'Sheet access denied. Confirm Editor access for sheets-writer-unstuck-dashboar@unstuck-dnc-and-merger.iam.gserviceaccount.com and check for protected rows or tabs, then Retry sync.')
        if response.status_code == 429:
            raise HTTPException(429, 'Google Sheets quota reached. The setup is saved; Retry sync shortly.')
        response.raise_for_status()
        return response.json()
    except (requests.RequestException, ValueError):
        raise HTTPException(503, 'Sheet sync could not be verified. The setup is saved. Retry sync; existing rows are checked before another row is added.') from None


def first(values, fallback=''):
    return values[0][0] if values and values[0] else fallback


def metadata_id(job_id):
    # Google enforces uniqueness within the sheet. The same setup's insert and
    # marker commit atomically, even if a paused writer outlives its DB lease.
    return (int(hashlib.sha256(job_id.encode()).hexdigest()[:8], 16) & 0x7fffffff) or 1


def location(gid, row):
    return {'sheet_id': gid, 'row': row, 'url': f'{SOURCE_URL}#gid={gid}&range=A{row}'}


def write(session, requests_batch, gid, title, row, guard=None):
    try:
        return call(session, 'post', BASE+':batchUpdate', guard=guard, json={'requests': requests_batch})
    except HTTPException as exc:
        raise HTTPException(exc.status_code, f'{title}, row {row}: {exc.detail} '+location(gid, row)['url'], headers=exc.headers) from None


def schema(session, section):
    if section not in TABS:
        raise HTTPException(400, 'Choose the authorised client or Biz Dev strategy-sheet section.')
    gid, title = TABS[section]
    metadata = call(session, 'get', BASE, params={'fields': 'sheets.properties'})
    tab = next((s['properties'] for s in metadata.get('sheets', []) if s['properties']['sheetId'] == gid), None)
    if not tab or tab['title'] != title:
        raise HTTPException(502, f'The {title} tab is missing or renamed. Restore it before Retry sync.')
    if not 2 <= tab['gridProperties']['rowCount'] <= 20000:
        raise HTTPException(502, f'Check the row count in {title}; supported range is 2–20,000 rows.')
    header_end = col(min(30, tab['gridProperties'].get('columnCount', 30))-1)
    headers = call(session, 'get', BASE+'/values/'+a1(title, 'A1:'+header_end+'1')).get('values', [[]])[0]
    names = [str(v).strip().lower() for v in headers]
    if sum(n.startswith('email') for n in names) != 1 or names.count('live') != 1 or names.count('order') != 1:
        raise HTTPException(502, f'Fix {title}, row 1: Email, Order and Live headings must each appear once. '+SOURCE_URL+f'#gid={gid}&range=A1')
    email = next(i for i, n in enumerate(names) if n.startswith('email'))
    if email != 0:
        raise HTTPException(502, f'Fix {title}, row 1: keep Email in column A before Retry sync.')
    operational = {'order', 'live', 'warming up', 'at risk', 'needs deleting', 'being replaced', 'client', 'biz dev lead'}
    columns = {n: i for i, n in enumerate(names) if n in operational}
    owner_key = 'client' if section == 'client' else 'biz dev lead'
    if owner_key not in columns:
        raise HTTPException(502, f'Fix {title}, row 1: the {owner_key.title()} column is missing. '+SOURCE_URL+f'#gid={gid}&range=A1')
    # Safe column names above define the write mask, not a guessed A:G layout.
    return gid, title, tab['gridProperties']['rowCount'], columns, owner_key


def publish(job, owner, guard=None):
    try:
        session = gs._authed_session()
    except Exception:
        raise HTTPException(503, 'Google Sheets credentials are not configured. The completed setup remains saved.') from None
    gid, title, row_count, columns, owner_key = schema(session, job['section'])
    marker = call(session, 'post', BASE+'/developerMetadata:search', json={'dataFilters': [{'developerMetadataLookup': {
        'metadataKey': 'gmail_setup_id', 'metadataValue': job['id'], 'visibility': 'DOCUMENT'}}]})
    matches = marker.get('matchedDeveloperMetadata', [])
    emails = call(session, 'get', BASE+'/values/'+a1(title, f'A2:A{row_count}'), params={'valueRenderOption': 'UNFORMATTED_VALUE'}).get('values', [])
    existing = [i+2 for i, row in enumerate(emails) if row and email_cell_identity(row[0]) == job['email'].lower()]
    if len(existing) > 1:
        rows = ', '.join(str(n) for n in existing)
        raise HTTPException(409, f'{title}: duplicate sending email in rows {rows}. Keep one intended row, then Retry sync.')
    marked = [m['developerMetadata']['location'].get('dimensionRange', {}) for m in matches]
    marked = [m for m in marked if m.get('sheetId') == gid and m.get('dimension') == 'ROWS']
    if matches and (len(marked) != 1 or not existing or marked[0]['startIndex']+1 != existing[0]):
        raise HTTPException(409, f'{title}: the previously synced row was changed. Check the setup email against its sheet row before Retry sync.')
    if existing:
        row = existing[0]
        ranges = [a1(title, f'A{row}')]+[a1(title, f'{col(index)}{row}') for index in columns.values()]
        current = call(session, 'get', BASE+'/values:batchGet', params=[('ranges', r) for r in ranges]+[('valueRenderOption', 'UNFORMATTED_VALUE')])
        current_ranges = current.get('valueRanges', [])
        if len(current_ranges) != len(columns)+1:
            raise HTTPException(502, f'{title}, row {row}: Google returned an incomplete row; Retry sync. '+location(gid, row)['url'])
        if email_cell_identity(first(current_ranges[0].get('values'))) != job['email'].lower():
            raise HTTPException(409, f'{title}, row {row}: the sending email changed. Check the intended row before retrying. '+location(gid, row)['url'])
        values = {name: first(item.get('values')) for name, item in zip(columns, current_ranges[1:])}
        for flag in ('at risk', 'needs deleting', 'being replaced'):
            if checked(values.get(flag)):
                raise HTTPException(409, f'{title}, row {row}: {flag.title()} is checked. Review this row before marking it Live. '+location(gid, row)['url'])
        if str(values.get(owner_key, '')).strip().casefold() not in ('', owner.strip().casefold()):
            raise HTTPException(409, f'{title}, row {row}: the owner/client differs from this setup. Correct the sheet or the setup, then Retry sync. '+location(gid, row)['url'])
        updates = []
        if values.get('order') != job['order']:
            updates.append({'updateCells': {'range': {'sheetId': gid, 'startRowIndex': row-1, 'endRowIndex': row, 'startColumnIndex': columns['order'], 'endColumnIndex': columns['order']+1}, 'rows': [{'values': [{'userEnteredValue': {'numberValue': job['order']}}]}], 'fields': 'userEnteredValue'}})
        if not checked(values.get('live')):
            updates.append({'updateCells': {'range': {'sheetId': gid, 'startRowIndex': row-1, 'endRowIndex': row, 'startColumnIndex': columns['live'], 'endColumnIndex': columns['live']+1}, 'rows': [{'values': [{'userEnteredValue': {'boolValue': True}}]}], 'fields': 'userEnteredValue'}})
        if not values.get(owner_key):
            updates.append({'updateCells': {'range': {'sheetId': gid, 'startRowIndex': row-1, 'endRowIndex': row, 'startColumnIndex': columns[owner_key], 'endColumnIndex': columns[owner_key]+1}, 'rows': [{'values': [{'userEnteredValue': {'stringValue': owner}}]}], 'fields': 'userEnteredValue'}})
        if updates:
            write(session, updates, gid, title, row, guard=guard)
    else:
        # Insert beneath the last email, preserving other rows and later recovery
        # columns. All insert/write/validation/marker requests form one transaction.
        display_email = job.get('sheet_display_email', job['email'])
        if email_cell_identity(display_email) != job['email'].lower():
            raise HTTPException(400, 'The sheet display email differs from this setup. Refresh before retrying the backup.')
        last = max((i+2 for i, r in enumerate(emails) if r and str(r[0]).strip()), default=1)
        index = last  # zero-based insertion point immediately after last email
        row = index+1
        requests_batch = []
        if index == row_count:
            requests_batch.append({'appendDimension': {'sheetId': gid, 'dimension': 'ROWS', 'length': 1}})
        else:
            requests_batch.append({'insertDimension': {'range': {'sheetId': gid, 'dimension': 'ROWS', 'startIndex': index, 'endIndex': index+1}, 'inheritFromBefore': True}})
        # Copy formatting only. Formula values, passwords and recovery data are
        # never cloned; existing rows shift intact with the native insertion.
        last_col = max(columns.values())+1
        requests_batch.append({'copyPaste': {'source': {'sheetId': gid, 'startRowIndex': index-1 if index>1 else index+1, 'endRowIndex': index if index>1 else index+2, 'startColumnIndex': 0, 'endColumnIndex': last_col}, 'destination': {'sheetId': gid, 'startRowIndex': index, 'endRowIndex': index+1, 'startColumnIndex': 0, 'endColumnIndex': last_col}, 'pasteType': 'PASTE_FORMAT'}})
        contents = {0: {'stringValue': display_email}, columns['order']: {'numberValue': job['order']}, columns[owner_key]: {'stringValue': owner}}
        for flag in ('live', 'warming up', 'at risk', 'needs deleting', 'being replaced'):
            if flag in columns:
                contents[columns[flag]] = {'boolValue': flag == 'live'}
                requests_batch.append({'setDataValidation': {'range': {'sheetId': gid, 'startRowIndex': index, 'endRowIndex': index+1, 'startColumnIndex': columns[flag], 'endColumnIndex': columns[flag]+1}, 'rule': {'condition': {'type': 'BOOLEAN'}, 'strict': True, 'showCustomUi': True}}})
        for column, value in contents.items():
            requests_batch.append({'updateCells': {'range': {'sheetId': gid, 'startRowIndex': index, 'endRowIndex': index+1, 'startColumnIndex': column, 'endColumnIndex': column+1}, 'rows': [{'values': [{'userEnteredValue': value}]}], 'fields': 'userEnteredValue'}})
        requests_batch.append({'createDeveloperMetadata': {'developerMetadata': {'metadataId': metadata_id(job['id']), 'metadataKey': 'gmail_setup_id', 'metadataValue': job['id'], 'visibility': 'DOCUMENT', 'location': {'dimensionRange': {'sheetId': gid, 'dimension': 'ROWS', 'startIndex': index, 'endIndex': index+1}}}}})
        write(session, requests_batch, gid, title, row, guard=guard)
    # Check the exact row after the write. A timeout leaves the setup retryable.
    checks = [a1(title, f'A{row}'), a1(title, f'{col(columns["live"])}{row}'), a1(title, f'{col(columns["order"])}{row}')]
    try:
        readback = call(session, 'get', BASE+'/values:batchGet', guard=guard, params=[('ranges', r) for r in checks]+[('valueRenderOption', 'UNFORMATTED_VALUE')]).get('valueRanges', [])
    except HTTPException as exc:
        raise HTTPException(exc.status_code, f'{title}, row {row}: {exc.detail} '+location(gid, row)['url'], headers=exc.headers) from None
    if len(readback) != 3 or email_cell_identity(first(readback[0].get('values'))) != job['email'].lower() or not checked(first(readback[1].get('values'), False)) or first(readback[2].get('values')) != job['order']:
        raise HTTPException(502, f'{title}, row {row}: could not verify the email and Live checkbox. Retry sync before editing other rows. '+location(gid, row)['url'])
    return {'tab': title, **location(gid, row)}


def deactivate(job, owner, guard=None):
    """Untick only the known sending email's Live cell; never create a row."""
    try:
        session = gs._authed_session()
    except Exception:
        raise HTTPException(503, 'Google Sheets credentials are not configured. The deleted Gmail remains saved.') from None
    gid, title, row_count, columns, owner_key = schema(session, job['section'])
    marker = call(session, 'post', BASE+'/developerMetadata:search', json={'dataFilters': [{'developerMetadataLookup': {
        'metadataKey': 'gmail_setup_id', 'metadataValue': job['id'], 'visibility': 'DOCUMENT'}}]})
    matches = marker.get('matchedDeveloperMetadata', [])
    emails = call(session, 'get', BASE+'/values/'+a1(title, f'A2:A{row_count}'), params={'valueRenderOption': 'UNFORMATTED_VALUE'}).get('values', [])
    email = str(job.get('email') or '').strip().lower()
    existing = [i+2 for i, values in enumerate(emails) if email and values and email_cell_identity(values[0]) == email]
    if len(existing) > 1:
        details = ' '.join(f'row {row}: '+location(gid, row)['url'] for row in existing)
        raise HTTPException(409, f'{title}: duplicate sending email. Keep one intended row before retrying. '+details)
    marked = [m.get('developerMetadata', {}).get('location', {}).get('dimensionRange', {})
              for m in matches if isinstance(m, dict)]
    marked = [m for m in marked if m.get('sheetId') == gid and m.get('dimension') == 'ROWS']
    if matches:
        consistent = len(matches) == len(marked) == 1
        target = marked[0] if len(marked) == 1 else {}
        start, end = target.get('startIndex'), target.get('endIndex')
        consistent = consistent and type(start) is int and type(end) is int and 1 <= start < row_count and end == start+1
        consistent = consistent and len(existing) == 1 and start+1 == existing[0]
        if not consistent:
            rows = sorted(set(existing + [m['startIndex']+1 for m in marked
                if type(m.get('startIndex')) is int and 1 <= m['startIndex'] < row_count]))
            details = ' '.join(f'row {row}: '+location(gid, row)['url'] for row in rows)
            raise HTTPException(409, f'{title}: the previously synced row was changed. Check the setup email against its marked sheet row before retrying. '+details)
    if not existing:
        return {'tab': title, 'sheet_id': gid, 'live': False, 'missing': True,
                'message': 'No matching sheet row was found. No row was added or changed.'}
    row = existing[0]
    row_location = location(gid, row)
    ranges = [a1(title, f'A{row}'), a1(title, f'{col(columns[owner_key])}{row}'), a1(title, f'{col(columns["live"])}{row}')]
    current = call(session, 'get', BASE+'/values:batchGet', params=[('ranges', value) for value in ranges]+[('valueRenderOption', 'UNFORMATTED_VALUE')]).get('valueRanges', [])
    if len(current) != 3:
        raise HTTPException(502, f'{title}, row {row}: Google returned an incomplete row; retry before changing Live. '+row_location['url'])
    if email_cell_identity(first(current[0].get('values'))) != email:
        raise HTTPException(409, f'{title}, row {row}: the sending email changed. Check the intended row before retrying. '+row_location['url'])
    if str(first(current[1].get('values'))).strip().casefold() not in ('', owner.strip().casefold()):
        raise HTTPException(409, f'{title}, row {row}: the owner/client differs from this setup. Correct the sheet or the setup before retrying. '+row_location['url'])
    if first(current[2].get('values'), None) is not False:
        updates = [{'updateCells': {'range': {'sheetId': gid, 'startRowIndex': row-1, 'endRowIndex': row,
            'startColumnIndex': columns['live'], 'endColumnIndex': columns['live']+1},
            'rows': [{'values': [{'userEnteredValue': {'boolValue': False}}]}], 'fields': 'userEnteredValue'}}]
        write(session, updates, gid, title, row, guard=guard)
    checks = [a1(title, f'A{row}'), a1(title, f'{col(columns["live"])}{row}')]
    try:
        readback = call(session, 'get', BASE+'/values:batchGet', guard=guard,
            params=[('ranges', value) for value in checks]+[('valueRenderOption', 'UNFORMATTED_VALUE')]).get('valueRanges', [])
    except HTTPException as exc:
        raise HTTPException(exc.status_code, f'{title}, row {row}: {exc.detail} '+row_location['url'], headers=exc.headers) from None
    if len(readback) != 2 or email_cell_identity(first(readback[0].get('values'))) != email or first(readback[1].get('values'), None) is not False:
        raise HTTPException(502, f'{title}, row {row}: could not verify the email and unchecked Live checkbox. Retry before editing other rows. '+row_location['url'])
    return {'tab': title, **row_location, 'live': False}


def readiness():
    try:
        session = gs._authed_session()
        r = call(session, 'get', f'https://www.googleapis.com/drive/v3/files/{SHEET_ID}', params={'fields': 'capabilities(canEdit)'})
        ready = r.get('capabilities', {}).get('canEdit') is True
        return {'ready': ready, 'message': 'Sheet write access ready.' if ready else 'Share the strategy sheet as Editor with sheets-writer-unstuck-dashboar@unstuck-dnc-and-merger.iam.gserviceaccount.com.'}
    except HTTPException as exc:
        return {'ready': False, 'message': str(exc.detail)}
    except Exception:
        return {'ready': False, 'message': 'Google Sheets credentials are not available.'}
