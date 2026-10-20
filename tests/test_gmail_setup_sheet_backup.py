"""Synthetic strategy-sheet backup and deleted-Gmail Live reconciliation."""
import copy
from uuid import uuid4

import pytest
from fastapi import HTTPException

from app.utils.gmail_setup import sheets
from test_gmail_setup import sheet_session


def tracked(section='client', **patch):
    return {'id': str(uuid4()), 'section': section, 'email': 'old@existing.com', 'order': 99, **patch}


def marker(session, job, row=2, **patch):
    value = {'metadataId': sheets.metadata_id(job['id']), 'metadataKey': 'gmail_setup_id',
             'metadataValue': job['id'], 'visibility': 'DOCUMENT', 'location': {'dimensionRange': {
                 'sheetId': session.gid, 'dimension': 'ROWS', 'startIndex': row-1, 'endIndex': row}}}
    value['location']['dimensionRange'].update(patch)
    session.markers.append(value)


def mutations(session):
    return [body for method, url, body in session.calls if method == 'post' and url == sheets.BASE+':batchUpdate']


@pytest.mark.parametrize('section', ['client', 'bizdev'])
@pytest.mark.parametrize('marked', [False, True])
def test_deactivate_unticks_only_live_on_exact_tab_and_retries_without_writing(sheet_session, section, marked):
    session = sheet_session(section)
    job = tracked(section)
    session.headers.extend(['Password', 'Recovery codes'])
    session.rows[1].extend(['synthetic-private-password', 'synthetic-recovery-code'])
    for flag in ('At risk', 'Needs deleting', 'Being replaced', 'Warming up'):
        if flag in session.headers:
            session.rows[1][session.headers.index(flag)] = True
    if marked:
        marker(session, job)
    before, markers = copy.deepcopy(session.rows), copy.deepcopy(session.markers)
    result = sheets.deactivate(job, 'Existing')
    expected = copy.deepcopy(before)
    expected[1][session.headers.index('Live')] = False
    assert session.rows == expected and session.markers == markers and session.inserts == 0
    assert result == {'tab': session.title, **sheets.location(session.gid, 2), 'live': False}
    batch = mutations(session)
    assert len(batch) == 1 and len(batch[0]['requests']) == 1
    update = batch[0]['requests'][0]['updateCells']
    assert update['range']['startColumnIndex'] == session.headers.index('Live')
    assert update['rows'] == [{'values': [{'userEnteredValue': {'boolValue': False}}]}]
    assert sheets.deactivate(job, ' existing ') == result and len(mutations(session)) == 1
    # The password and recovery columns are never requested or echoed.
    assert 'synthetic-private-password' not in str(session.calls)+str(result)
    assert 'synthetic-recovery-code' not in str(session.calls)+str(result)


@pytest.mark.parametrize('section', ['client', 'bizdev'])
def test_deactivate_missing_email_never_appends_or_changes_another_row(sheet_session, section):
    session = sheet_session(section)
    before = copy.deepcopy(session.rows)
    result = sheets.deactivate(tracked(section, email='absent@example.com'), 'Different owner')
    assert result['missing'] is True and result['live'] is False
    assert result['sheet_id'] == session.gid and result['message'].startswith('No matching sheet row')
    assert 'row' not in result and session.rows == before and not mutations(session) and session.inserts == 0


def test_deactivate_blank_planning_email_never_matches_a_formatted_blank_row(sheet_session):
    session = sheet_session()
    session.rows.append(['', 7, True, False, False, False, 'Existing'])
    before = copy.deepcopy(session.rows)
    assert sheets.deactivate(tracked(email=''), 'Existing')['missing'] is True
    assert session.rows == before and not mutations(session)


def test_deactivate_lost_write_response_reconciles_committed_unchecked_value_without_second_write(sheet_session):
    session = sheet_session()
    job = tracked()
    marker(session, job)
    session.lose_response = True
    with pytest.raises(HTTPException) as exc:
        sheets.deactivate(job, 'Existing')
    assert exc.value.status_code == 503 and 'row 2' in exc.value.detail and 'range=A2' in exc.value.detail
    assert session.rows[1][session.headers.index('Live')] is False and len(mutations(session)) == 1
    assert sheets.deactivate(job, 'Existing')['live'] is False
    assert len(mutations(session)) == 1 and session.inserts == 0


def test_deactivate_owner_conflict_reports_exact_row_without_writing(sheet_session):
    session = sheet_session()
    before = copy.deepcopy(session.rows)
    with pytest.raises(HTTPException) as exc:
        sheets.deactivate(tracked(), 'Another client')
    assert exc.value.status_code == 409 and 'owner/client differs' in exc.value.detail
    assert sheets.location(session.gid, 2)['url'] in exc.value.detail
    assert session.rows == before and not mutations(session)


@pytest.mark.parametrize('section', ['client', 'bizdev'])
def test_deactivate_duplicate_email_including_master_label_requires_manual_row_review(sheet_session, section):
    session = sheet_session(section)
    duplicate = copy.deepcopy(session.rows[1])
    duplicate[0] += ' (Master)'
    session.rows.append(duplicate)
    before = copy.deepcopy(session.rows)
    with pytest.raises(HTTPException) as exc:
        sheets.deactivate(tracked(section), 'Existing')
    assert exc.value.status_code == 409 and 'duplicate sending email' in exc.value.detail
    assert all(sheets.location(session.gid, row)['url'] in exc.value.detail for row in (2, 3))
    assert session.rows == before and not mutations(session)


@pytest.mark.parametrize('problem', ['moved', 'other_tab', 'multiple', 'invalid_range', 'email_missing'])
def test_deactivate_rejects_changed_or_ambiguous_metadata_before_mutation(sheet_session, problem):
    session = sheet_session()
    job = tracked()
    session.rows.append(['unrelated@example.com', 7, True, False, False, False, 'Unrelated'])
    marker(session, job, row=3 if problem == 'moved' else 2)
    if problem == 'other_tab':
        session.markers[0]['location']['dimensionRange']['sheetId'] = sheets.TABS['bizdev'][0]
    elif problem == 'multiple':
        marker(session, job, row=3)
    elif problem == 'invalid_range':
        session.markers[0]['location']['dimensionRange']['endIndex'] = 4
    elif problem == 'email_missing':
        session.rows[1][0] = 'changed@example.com'
    before = copy.deepcopy(session.rows)
    with pytest.raises(HTTPException) as exc:
        sheets.deactivate(job, 'Existing')
    assert exc.value.status_code == 409 and 'previously synced row was changed' in exc.value.detail
    if problem != 'other_tab':
        assert 'range=A2' in exc.value.detail
    assert session.rows == before and not mutations(session)


@pytest.mark.parametrize('bad_value', ['FALSE', '', 0, True])
def test_deactivate_requires_typed_false_readback_and_points_to_row(sheet_session, bad_value):
    session = sheet_session()
    original = session.post
    def post(url, json, **kwargs):
        result = original(url, json, **kwargs)
        if url == sheets.BASE+':batchUpdate':
            session.rows[1][session.headers.index('Live')] = bad_value
        return result
    session.post = post
    with pytest.raises(HTTPException) as exc:
        sheets.deactivate(tracked(), 'Existing')
    assert exc.value.status_code == 502 and 'unchecked Live checkbox' in exc.value.detail and 'range=A2' in exc.value.detail
    assert len(mutations(session)) == 1 and session.inserts == 0


def test_deactivate_rechecks_exact_email_before_touching_live(sheet_session):
    session = sheet_session()
    original = session.get
    def get(url, params=None, **kwargs):
        if url.endswith('/values:batchGet'):
            session.rows[1][0] = 'someone-else@example.com'
        return original(url, params=params, **kwargs)
    session.get = get
    with pytest.raises(HTTPException, match='sending email changed'):
        sheets.deactivate(tracked(), 'Existing')
    assert session.rows[1][session.headers.index('Live')] is True and not mutations(session)


def test_publish_rechecks_exact_email_after_scan_before_updating_existing_row(sheet_session):
    session = sheet_session()
    original = session.get
    def get(url, params=None, **kwargs):
        if url.endswith('/values:batchGet'):
            session.rows[1][0] = 'another-client@example.com'
        return original(url, params=params, **kwargs)
    session.get = get
    before = copy.deepcopy(session.rows[1][1:])
    with pytest.raises(HTTPException) as exc:
        sheets.publish(tracked(), 'Existing')
    assert exc.value.status_code == 409 and 'sending email changed' in exc.value.detail
    assert sheets.location(session.gid, 2)['url'] in exc.value.detail
    assert session.rows[1][1:] == before and not mutations(session) and session.inserts == 0


@pytest.mark.parametrize('operation', ['insert', 'update', 'deactivate'])
def test_guard_can_stop_every_batch_write_after_limiter_without_side_effects(sheet_session, operation, monkeypatch):
    session = sheet_session()
    job = tracked(email='new@example.com') if operation == 'insert' else tracked()
    before = copy.deepcopy(session.rows)
    limited = []
    monkeypatch.setattr(sheets.gs._write_limiter, 'acquire', lambda: limited.append(True))
    def guard():
        assert limited
        raise HTTPException(409, 'Stale operation')
    function = sheets.deactivate if operation == 'deactivate' else sheets.publish
    with pytest.raises(HTTPException, match='Stale operation'):
        function(job, 'Existing', guard=guard)
    assert session.rows == before and not mutations(session) and session.inserts == 0


@pytest.mark.parametrize('operation', ['publish', 'deactivate'])
def test_guard_also_protects_readback_when_no_write_is_needed(sheet_session, operation):
    session = sheet_session()
    job = tracked(order=1)
    if operation == 'deactivate':
        session.rows[1][session.headers.index('Live')] = False
    def guard():
        raise HTTPException(409, 'Stale readback')
    with pytest.raises(HTTPException, match='Stale readback'):
        getattr(sheets, operation)(job, 'Existing', guard=guard)
    assert not mutations(session)


@pytest.mark.parametrize('suffix', [' (Master)', ' (Secondary: masterdomain.com is Master)'])
def test_publish_inserts_label_once_then_reconciles_and_deactivates_by_canonical_email(sheet_session, suffix):
    session = sheet_session()
    job = tracked(email='new@example.com', order=2, sheet_display_email='new@example.com'+suffix)
    result = sheets.publish(job, 'New client')
    assert session.rows[2][0] == job['sheet_display_email'] and session.inserts == 1
    assert sheets.publish({**job, 'sheet_display_email': job['email']}, 'New client') == result
    assert session.rows[2][0] == job['sheet_display_email'] and session.inserts == 1
    assert sheets.deactivate(job, 'New client')['live'] is False
    assert session.rows[2][0] == job['sheet_display_email'] and session.inserts == 1


@pytest.mark.parametrize('suffix', [' (Master)', ' (Secondary: masterdomain.com is Master)'])
def test_publishing_existing_plain_email_never_relabels_column_a(sheet_session, suffix):
    session = sheet_session()
    job = tracked(order=1, sheet_display_email='old@existing.com'+suffix)
    sheets.publish(job, 'Existing')
    assert session.rows[1][0] == job['email'] and not mutations(session) and session.inserts == 0


def test_mismatched_display_label_cannot_insert_another_sending_email(sheet_session):
    session = sheet_session()
    job = tracked(email='new@example.com', sheet_display_email='another@example.com (Master)')
    before = copy.deepcopy(session.rows)
    with pytest.raises(HTTPException, match='display email differs'):
        sheets.publish(job, 'Existing')
    assert session.rows == before and not mutations(session) and session.inserts == 0
