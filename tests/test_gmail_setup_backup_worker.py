"""Durable intent races and merged results, with synthetic sheet rows only."""
import copy
import json
from datetime import datetime, timedelta, timezone

import pytest
from fastapi import HTTPException

from app.utils.gmail_setup import backup, service as s, sheets, store
from test_gmail_setup import board_db, USER, OTHER, people, change, create_setup, devices

_RESUME_WORKER = backup.resume # board_db disables HTTP scheduling, not these explicit worker checks.


def job_at(job_id, user=USER):
    return next(job for job in s.board(user)['state']['setups'] if job['id'] == job_id)


def stage(job_id, value):
    return change('setup_stage', {'id': job_id, 'stage': value,
        'confirmed_auto_renew_stopped': value == s.GMAIL_DELETED})


class SyntheticSheet:
    def __init__(self):
        self.rows = {}
        self.calls = []
        self.writes = []
        self.hooks = {}

    def hook(self, operation, point):
        callback = self.hooks.pop((operation, point), None)
        if callback:
            callback()

    def location(self, job):
        row = self.rows[job['id']]
        return {'tab': sheets.TABS[job['section']][1], **sheets.location(sheets.TABS[job['section']][0], row['row'])}

    def publish(self, job, owner, guard=None):
        self.calls.append(('activate', copy.deepcopy(job), owner))
        self.hook('activate', 'before')
        assert guard is not None
        guard()
        row = self.rows.setdefault(job['id'], {'row': 2+len(self.rows), 'email': job.get('sheet_display_email', job['email']),
            'owner': owner, 'order': job['order'], 'recovery': 'synthetic-retained-recovery', 'live': False})
        row['live'] = True
        self.writes.append(('activate', job['id']))
        self.hook('activate', 'after')
        guard()
        return self.location(job)

    def deactivate(self, job, owner, guard=None):
        self.calls.append(('deactivate', copy.deepcopy(job), owner))
        self.hook('deactivate', 'before')
        assert guard is not None
        guard()
        if job['id'] not in self.rows:
            return {'tab': sheets.TABS[job['section']][1], 'sheet_id': sheets.TABS[job['section']][0],
                'live': False, 'missing': True, 'message': 'No matching sheet row was found. No row was added or changed.'}
        self.rows[job['id']]['live'] = False
        self.writes.append(('deactivate', job['id']))
        self.hook('deactivate', 'after')
        guard()
        return {**self.location(job), 'live': False}


@pytest.fixture
def synthetic_sheet(monkeypatch):
    provider = SyntheticSheet()
    monkeypatch.setattr(sheets, 'publish', provider.publish)
    monkeypatch.setattr(sheets, 'deactivate', provider.deactivate)
    monkeypatch.setattr(sheets.gs, '_authed_session', lambda: pytest.fail('Worker tests cannot contact Google'))
    return provider


def assert_synced_intent(job, live):
    assert job['sync']['status'] == 'synced'
    assert job['sync']['generation'] == job['sheet_intent']['generation']
    assert job['sync']['live'] is job['sheet_intent']['live'] is live
    assert not job['sync'].get('in_progress')


def test_stage_saves_queue_durable_intents_without_provider_io_and_drain_is_idempotent(board_db, synthetic_sheet):
    _, job_id = create_setup(people()['Dylan'])
    completed = stage(job_id, 'Setup complete')
    job = completed['state']['setups'][0]
    generation = job['sheet_intent']['generation']
    assert job['sync']['status'] == 'pending' and job['sheet_intent']['live'] is True
    assert not synthetic_sheet.calls and not synthetic_sheet.rows
    assert stage(job_id, 'Setup complete')['state']['setups'][0]['sheet_intent']['generation'] == generation
    backup.drain(OTHER)
    assert_synced_intent(job_at(job_id), True)
    assert synthetic_sheet.rows[job_id]['live'] is True and len(synthetic_sheet.rows) == 1
    writes = copy.deepcopy(synthetic_sheet.writes)
    backup.drain(USER)
    backup.sync_once(OTHER, job_id)
    assert synthetic_sheet.writes == writes and s.board(USER)['state']['sync_lock'] is None


@pytest.mark.parametrize('point', ['before', 'after'])
def test_complete_to_deleted_during_publish_converges_to_latest_false_generation(board_db, synthetic_sheet, point):
    _, job_id = create_setup(people()['Dylan'])
    stage(job_id, 'Setup complete')
    old_generation = job_at(job_id)['sheet_intent']['generation']
    observed = {}
    def delete_while_paused():
        lock = copy.deepcopy(s.board(USER)['state']['sync_lock'])
        assert lock['kind'] == 'sheet' and lock['generation'] == old_generation
        changed = stage(job_id, s.GMAIL_DELETED)
        assert changed['state']['sync_lock'] == lock
        observed['generation'] = job_at(job_id)['sheet_intent']['generation']
        assert observed['generation'] != old_generation
    synthetic_sheet.hooks[('activate', point)] = delete_while_paused
    backup.sync_once(USER, job_id)
    pending = job_at(job_id)
    assert pending['stage'] == s.GMAIL_DELETED and pending['sync']['status'] == 'pending'
    assert pending['sync']['generation'] == observed['generation']
    assert s.board(USER)['state']['sync_lock'] is None
    backup.drain(OTHER)
    final = job_at(job_id)
    assert_synced_intent(final, False)
    assert final['stage'] == s.GMAIL_DELETED and final['sync']['generation'] != old_generation
    if point == 'before':
        assert job_id not in synthetic_sheet.rows and not synthetic_sheet.writes
        assert final['sync']['missing'] is True
    else:
        assert synthetic_sheet.rows[job_id]['live'] is False
        assert synthetic_sheet.writes == [('activate', job_id), ('deactivate', job_id)]


def test_deleted_to_not_started_keeps_false_intent_and_does_not_restore_live(board_db, synthetic_sheet):
    _, job_id = create_setup(people()['Dylan'])
    stage(job_id, 'Setup complete');backup.drain(USER)
    original = copy.deepcopy(synthetic_sheet.rows[job_id])
    stage(job_id, s.GMAIL_DELETED)
    generation = job_at(job_id)['sheet_intent']['generation']
    def reopen():
        lock = copy.deepcopy(s.board(USER)['state']['sync_lock'])
        reopened = stage(job_id, 'Not started')
        assert reopened['state']['sync_lock'] == lock
        assert job_at(job_id)['sheet_intent']['generation'] == generation
    synthetic_sheet.hooks[('deactivate', 'before')] = reopen
    backup.drain(OTHER)
    assert job_at(job_id)['stage'] == 'Not started'
    assert_synced_intent(job_at(job_id), False)
    assert synthetic_sheet.rows[job_id] == {**original, 'live': False}


def test_complete_deleted_complete_race_keeps_only_the_latest_live_intent(board_db, synthetic_sheet):
    _, job_id = create_setup(people()['Dylan'])
    stage(job_id, 'Setup complete')
    generations = [job_at(job_id)['sheet_intent']['generation']]
    def replace_intent_twice():
        stage(job_id, s.GMAIL_DELETED)
        generations.append(job_at(job_id)['sheet_intent']['generation'])
        stage(job_id, 'Not started')
        stage(job_id, 'Setup complete')
        generations.append(job_at(job_id)['sheet_intent']['generation'])
    synthetic_sheet.hooks[('activate', 'before')] = replace_intent_twice
    backup.drain(USER)
    final = job_at(job_id)
    assert len(set(generations)) == 3 and final['sync']['generation'] == generations[-1]
    assert final['stage'] == 'Setup complete'
    assert_synced_intent(final, True)
    assert synthetic_sheet.writes == [('activate', job_id)] and len(synthetic_sheet.rows) == 1


@pytest.mark.parametrize('section', ['client', 'bizdev'])
def test_unrelated_edits_and_reopened_owner_device_changes_survive_a_slow_backup(board_db, synthetic_sheet, section):
    people_ids = people()
    first_phone = devices(people_ids['Dylan'])['state']['devices'][0]
    replacement = devices(people_ids['Matthew'])['state']['devices'][-1]
    record, job_id = create_setup(people_ids['Dylan'], first_phone['id'], section=section)
    other = change('client_save', {'name': 'Other client', 'account_manager_id': people_ids['Robyn']})['state']['clients'][-1]
    stage(job_id, 'Setup complete')
    generation = job_at(job_id)['sheet_intent']['generation']
    def edit_while_paused():
        lock = copy.deepcopy(s.board(USER)['state']['sync_lock'])
        actions = [
            ('setup_stage', {'id': job_id, 'stage': 'Not started'}),
            ('setup_assign', {'id': job_id, 'operator_id': people_ids['Matthew'], 'device_id': replacement['id']}),
            ('setup_notes', {'id': job_id, 'notes': 'Saved while the sheet provider was paused'}),
            ('device_update', {'id': replacement['id'], 'asset_tag': 'SYNTHETIC-REPLACEMENT'}),
            ('client_save', {'id': other['id'], 'name': 'Other client renamed', 'account_manager_id': people_ids['Dylan']})]
        for action, data in actions:
            updated = change(action, data, user=OTHER)
            assert updated['state']['sync_lock'] == lock
        assert job_at(job_id)['sheet_intent']['generation'] == generation
    synthetic_sheet.hooks[('activate', 'before')] = edit_while_paused
    backup.drain(USER)
    final = job_at(job_id)
    assert final['stage'] == 'Not started' and final['operator_id'] == people_ids['Matthew']
    assert final['device_id'] == replacement['id'] and final['notes'] == 'Saved while the sheet provider was paused'
    assert final['purchase_device']['id'] == first_phone['id'] and final['order'] == 1
    assert_synced_intent(final, True)
    state = s.board(USER)['state']
    assert next(row for row in state['clients'] if row['id'] == other['id'])['name'] == 'Other client renamed'
    assert next(row for row in state['devices'] if row['id'] == replacement['id'])['asset_tag'] == 'SYNTHETIC-REPLACEMENT'
    assert next(row for row in state['devices'] if row['id'] == first_phone['id'])['use_count'] == 1
    assert synthetic_sheet.rows[job_id]['owner'] == ('Cedar client' if section == 'client' else 'Dylan')
    assert state['sync_lock'] is None


def test_final_cas_retry_merges_an_edit_instead_of_replacing_the_latest_board(board_db, synthetic_sheet, monkeypatch):
    _, job_id = create_setup(people()['Dylan'])
    stage(job_id, 'Setup complete')
    original = store.save
    raced = []
    def save(user, state, revision):
        if not raced and state['sync_lock'] is None and state['events'][-1]['action'] == 'sheet_backup_synced':
            raced.append(True)
            change('setup_notes', {'id': job_id, 'notes': 'Concurrent final-save edit'}, user=OTHER)
            raise HTTPException(409, 'Synthetic CAS conflict')
        return original(user, state, revision)
    monkeypatch.setattr(store, 'save', save)
    backup.sync_once(USER, job_id)
    assert raced and job_at(job_id)['notes'] == 'Concurrent final-save edit'
    assert_synced_intent(job_at(job_id), True)
    assert synthetic_sheet.writes == [('activate', job_id)] and s.board(USER)['state']['sync_lock'] is None


def test_missing_deactivation_discards_an_obsolete_previously_synced_row_link(board_db, synthetic_sheet):
    _, job_id = create_setup(people()['Dylan'])
    stage(job_id, 'Setup complete');backup.drain(USER)
    assert job_at(job_id)['sync']['row'] == 2 and job_at(job_id)['sync']['url']
    synthetic_sheet.rows.clear() # A human removed the original strategy-sheet row.
    stage(job_id, s.GMAIL_DELETED)
    assert job_at(job_id)['sync']['row'] == 2
    backup.sync_once(USER, job_id)
    final = job_at(job_id)
    assert_synced_intent(final, False)
    assert final['sync']['missing'] is True and 'row' not in final['sync'] and 'url' not in final['sync']
    assert not synthetic_sheet.rows


def test_retry_after_lost_provider_ack_reconciles_one_row_with_the_original_identity(board_db, synthetic_sheet):
    _, job_id = create_setup(people()['Dylan'])
    stage(job_id, 'Setup complete')
    generation = job_at(job_id)['sheet_intent']['generation']
    def lost_ack():
        raise HTTPException(503, 'Synthetic response lost after commit')
    synthetic_sheet.hooks[('activate', 'after')] = lost_ack
    backup.drain(USER)
    failed = job_at(job_id)
    assert failed['sync']['status'] == 'error' and failed['stage'] == 'Setup complete'
    assert failed['sync']['generation'] == generation and len(synthetic_sheet.rows) == 1
    original = copy.deepcopy(synthetic_sheet.rows[job_id])
    retried = backup.queue_retry(OTHER, job_id)
    assert retried['state']['setups'][0]['sync']['status'] == 'pending'
    backup.drain(OTHER)
    assert_synced_intent(job_at(job_id), True)
    assert job_at(job_id)['sync']['generation'] != generation
    assert synthetic_sheet.rows[job_id] == original and len(synthetic_sheet.rows) == 1


@pytest.mark.parametrize('kind', ['dns_setup', 'dmarc'])
def test_dns_claim_still_blocks_edits_and_sheet_backups_without_provider_calls(board_db, synthetic_sheet, kind):
    record, job_id = create_setup(people()['Dylan'])
    client_id = record['state']['setups'][0]['client_id']
    stage(job_id, 'Setup complete')
    record = s.board(USER)
    record['state']['sync_lock'] = {'kind': kind, 'token': 'synthetic-dns-owner', 'job_id': job_id, 'at': s.now()}
    store.save(USER, record['state'], record['revision'])
    before = s.board(USER)
    calls = [lambda: change('setup_notes', {'id': job_id, 'notes': 'Cannot edit yet'}),
        lambda: backup.sync_once(USER, job_id), lambda: backup.queue_retry(USER, job_id),
        lambda: backup.queue_client(USER, client_id, before['revision'])]
    for call in calls:
        with pytest.raises(HTTPException) as exc:
            call()
        assert exc.value.status_code == 409 and s.board(USER) == before
    backup.drain(USER)
    assert s.board(USER) == before and not synthetic_sheet.calls


def test_other_workspace_cannot_resolve_a_queued_job_or_reach_the_provider(board_db, synthetic_sheet):
    record, job_id = create_setup(people()['Dylan'])
    client_id = record['state']['setups'][0]['client_id']
    stage(job_id, 'Setup complete')
    before = s.board(USER)
    stranger = {'email': 'member@another-workspace.test', 'role': 'admin', 'tools': ['gmail_set']}
    for call in (lambda: backup.sync_once(stranger, job_id), lambda: backup.queue_retry(stranger, job_id),
                 lambda: backup.queue_client(stranger, client_id, 0)):
        with pytest.raises(HTTPException) as exc:
            call()
        assert exc.value.status_code == 404
    backup.drain(stranger)
    assert s.board(USER) == before and not synthetic_sheet.calls and len(board_db) == 1


def test_restart_resumes_an_expired_durable_claim_without_manual_retry_or_new_generation(board_db, synthetic_sheet, monkeypatch):
    _, job_id = create_setup(people()['Dylan'])
    stage(job_id, 'Setup complete')
    record = s.board(USER)
    generation = job_at(job_id)['sheet_intent']['generation']
    record['state']['sync_lock'] = {'kind': 'sheet', 'token': 'dead-process-claim', 'job_id': job_id,
        'generation': generation, 'at': (datetime.now(timezone.utc)-timedelta(seconds=s.LEASE_SECONDS+1)).isoformat()}
    record['state']['setups'][0]['sync']['in_progress'] = True
    store.save(USER, record['state'], record['revision'])
    monkeypatch.setattr(backup, '_running', set()) # A fresh process has no in-memory worker state.
    _RESUME_WORKER(OTHER)
    final = job_at(job_id)
    assert_synced_intent(final, True)
    assert final['sync']['generation'] == generation and synthetic_sheet.writes == [('activate', job_id)]
    assert s.board(USER)['state']['sync_lock'] is None and store.workspace(USER) not in backup._running
    _RESUME_WORKER(USER)
    assert synthetic_sheet.writes == [('activate', job_id)] and len(synthetic_sheet.rows) == 1


def test_guard_rejects_a_nearly_expired_claim_then_resume_retries_the_same_intent(board_db, synthetic_sheet):
    _, job_id = create_setup(people()['Dylan'])
    stage(job_id, 'Setup complete')
    generation = job_at(job_id)['sheet_intent']['generation']
    def expire_before_write():
        record = s.board(USER)
        record['state']['sync_lock']['at'] = (datetime.now(timezone.utc)-timedelta(seconds=s.LEASE_SECONDS-24)).isoformat()
        store.save(USER, record['state'], record['revision'])
    synthetic_sheet.hooks[('activate', 'before')] = expire_before_write
    backup.sync_once(USER, job_id)
    assert not synthetic_sheet.writes and not synthetic_sheet.rows
    assert job_at(job_id)['sync']['status'] == 'pending' and job_at(job_id)['sheet_intent']['generation'] == generation
    assert s.board(USER)['state']['sync_lock'] is None
    _RESUME_WORKER(OTHER)
    assert_synced_intent(job_at(job_id), True)
    assert job_at(job_id)['sync']['generation'] == generation and synthetic_sheet.writes == [('activate', job_id)]


def test_expired_worker_cannot_overwrite_a_replacement_workers_success(board_db, synthetic_sheet):
    _, job_id = create_setup(people()['Dylan'])
    stage(job_id, 'Setup complete')
    generation = job_at(job_id)['sheet_intent']['generation']
    replacement = {}
    def resume_in_another_process():
        record = s.board(USER)
        replacement['old_token'] = record['state']['sync_lock']['token']
        record['state']['sync_lock']['at'] = (datetime.now(timezone.utc)-timedelta(seconds=s.LEASE_SECONDS+1)).isoformat()
        store.save(USER, record['state'], record['revision'])
        backup.sync_once(OTHER, job_id)
        replacement['record'] = s.board(USER)
    synthetic_sheet.hooks[('activate', 'before')] = resume_in_another_process
    returned = backup.sync_once(USER, job_id)
    assert_synced_intent(job_at(job_id), True)
    assert returned == replacement['record'] == s.board(USER)
    assert job_at(job_id)['sync']['generation'] == generation
    assert synthetic_sheet.writes == [('activate', job_id)] and len(synthetic_sheet.rows) == 1
    assert s.board(USER)['state']['sync_lock'] is None


def test_public_board_redacts_private_intents_identity_and_lease_without_mutating_storage(board_db, synthetic_sheet):
    _, active = create_setup(people()['Dylan'])
    _, archived = create_setup(people()['Dylan'], domain='archivedbackup.com', email='me@archivedbackup.com')
    _, planning = create_setup(people()['Dylan'], domain='planningbackup.com', email='me@planningbackup.com')
    stage(active, 'Setup complete');stage(archived, 'Setup complete')
    change('setup_remove', {'id': archived})
    record = s.board(USER)
    record['state']['sync_lock'] = {'kind': 'sheet', 'token': 'SYNTHETIC-PRIVATE-LEASE', 'job_id': active,
        'generation': job_at(active)['sheet_intent']['generation'], 'at': s.now()}
    for job in record['state']['setups']:
        if job['id'] in (active, archived):
            job['sheet_identity']['private_test'] = 'SYNTHETIC-PRIVATE-IDENTITY'
            job['sheet_intent']['snapshot']['private_test'] = 'SYNTHETIC-PRIVATE-INTENT'
    stored = store.save(USER, record['state'], record['revision'])
    before = copy.deepcopy(stored)
    public = s.result(stored)
    public_jobs = {job['id']: job for job in public['state']['setups']}
    assert 'sync_lock' not in public['state']
    assert all('sheet_identity' not in job and 'sheet_intent' not in job for job in public_jobs.values())
    assert public_jobs[active]['sheet_identity_locked'] is public_jobs[archived]['sheet_identity_locked'] is True
    assert public_jobs[planning]['sheet_identity_locked'] is False
    assert public_jobs[active]['sync']['status'] == 'pending'
    assert all(value not in json.dumps(public) for value in ('SYNTHETIC-PRIVATE-LEASE', 'SYNTHETIC-PRIVATE-IDENTITY', 'SYNTHETIC-PRIVATE-INTENT'))
    assert stored == before == s.board(USER) and not synthetic_sheet.calls
