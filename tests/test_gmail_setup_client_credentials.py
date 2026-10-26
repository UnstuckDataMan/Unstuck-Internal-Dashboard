"""Atomic optional client credentials, using in-memory storage and no providers."""
import copy
import json

import pytest
from fastapi import HTTPException
from pydantic import ValidationError

from app import auth
from app.routers.gmail_setup import Client
from app.utils.domain_cluster import provisioning
from app.utils.gmail_setup import service as s, sheets, store
from test_gmail_setup import board_db, password_key, USER, OTHER, people, change, create_setup, devices


@pytest.fixture(autouse=True)
def no_external_provider(monkeypatch):
    def forbidden(*args, **kwargs):
        pytest.fail('Client credentials must never contact Sheets or Porkbun')
    monkeypatch.setattr(sheets, 'publish', forbidden)
    monkeypatch.setattr(sheets, 'deactivate', forbidden)
    monkeypatch.setattr(sheets.gs, '_authed_session', forbidden)
    monkeypatch.setattr(provisioning, 'Porkbun', forbidden)


def client_data(name='Credential client', **patch):
    return {'name': name, 'section': 'client', 'account_manager_id': people()['Robyn'], **patch}


def saved_client(record, identifier=None):
    return next(client for client in record['state']['clients'] if identifier is None or client['id'] == identifier)


def test_create_and_edit_optional_credentials_encrypt_password_and_preserve_card_leading_zeroes(board_db, password_key):
    first = '  Synthetic client credential 🔑\t '
    record = change('client_save', client_data(password=first, card_last4='0012'))
    client = saved_client(record)
    assert client['card_last4'] == '0012' and client['password_ciphertext'] != first
    assert client['password_updated_at'] and first not in json.dumps(board_db[store.workspace(USER)])
    assert 'password' not in client and s.reveal_client_password(OTHER, client['id']) == first
    public = s.result(record)['state']['clients'][0]
    assert public['card_last4'] == '0012' and public['has_password'] is True
    assert 'password_ciphertext' not in public and first not in json.dumps(s.result(record))
    previous_cipher = client['password_ciphertext']
    second = 'Replacement synthetic credential'
    edited = change('client_save', client_data(id=client['id'], name='Renamed credential client', password=second, card_last4='0000'))
    current = saved_client(edited)
    assert current['name'] == 'Renamed credential client' and current['card_last4'] == '0000'
    assert current['password_ciphertext'] != previous_cipher and s.reveal_client_password(USER, current['id']) == second
    assert second not in json.dumps(board_db[store.workspace(USER)])
    assert len(edited['state']['clients']) == 1 and edited['revision'] == record['revision']+1
    assert all(set(event) == {'at', 'by', 'action'} for event in edited['state']['events'])


@pytest.mark.parametrize('patch', [{}, {'password': None}, {'password': ''}])
def test_blank_optional_password_preserves_existing_credential_even_without_encryption_key(board_db, password_key, monkeypatch, patch):
    record = change('client_save', client_data(password='Existing synthetic secret', card_last4='0034'))
    original = copy.deepcopy(saved_client(record))
    monkeypatch.delenv('GMAIL_SETUP_PASSWORD_KEY', raising=False)
    monkeypatch.delenv('SESSION_SECRET', raising=False)
    edited = change('client_save', client_data(id=original['id'], name='Metadata only edit', **patch))
    current = saved_client(edited)
    assert current['password_ciphertext'] == original['password_ciphertext']
    assert current['password_updated_at'] == original['password_updated_at'] and current['card_last4'] == '0034'
    assert current['name'] == 'Metadata only edit'


@pytest.mark.parametrize('patch', [{}, {'password': None}, {'password': ''}])
def test_client_can_be_created_without_optional_password_or_configured_key(board_db, monkeypatch, patch):
    monkeypatch.delenv('GMAIL_SETUP_PASSWORD_KEY', raising=False)
    monkeypatch.delenv('SESSION_SECRET', raising=False)
    record = change('client_save', client_data(card_last4='0099', **patch))
    client = saved_client(record)
    assert client['card_last4'] == '0099' and not client.get('password_ciphertext')
    assert 'password_updated_at' not in client and s.result(record)['state']['clients'][0]['has_password'] is False


@pytest.mark.parametrize('password', [' ', '\t\n', 'Ünicode 🔐 password', 'x'*1024])
def test_optional_password_keeps_nonempty_whitespace_unicode_and_maximum_length(board_db, password_key, password):
    record = change('client_save', client_data(password=password))
    assert s.reveal_client_password(USER, saved_client(record)['id']) == password


@pytest.mark.parametrize('clear', ['', None])
def test_optional_card_can_be_cleared_without_removing_password(board_db, password_key, clear):
    record = change('client_save', client_data(password='Keep synthetic password', card_last4='0123'))
    client = saved_client(record)
    edited = change('client_save', client_data(id=client['id'], card_last4=clear))
    assert not saved_client(edited).get('card_last4')
    assert s.reveal_client_password(USER, client['id']) == 'Keep synthetic password'
    assert saved_client(edited)['password_ciphertext'] == client['password_ciphertext']
    removed = change('client_password', {'id': client['id'], 'password': None})
    assert s.result(removed)['state']['clients'][0]['has_password'] is False
    with pytest.raises(HTTPException) as exc:
        s.reveal_client_password(USER, client['id'])
    assert exc.value.status_code == 404


INVALID_CARDS = ['123', '12345', '4242424242424242', '１２３４', '١٢٣٤', '12٣4', '1234 ', ' 1234', '123\n', 1234, True, False, 12.34, [], {}]
INVALID_PASSWORDS = ['x'*1025, 1234, True, False, [], {}]


@pytest.mark.parametrize('field,value', [('card_last4', value) for value in INVALID_CARDS]+[('password', value) for value in INVALID_PASSWORDS])
def test_invalid_credentials_are_rejected_without_metadata_or_credential_changes(board_db, password_key, field, value):
    original = change('client_save', client_data(password='Original synthetic credential', card_last4='0010', sender_full_name='Original Sender'))
    client = saved_client(original)
    patch = client_data(id=client['id'], name='Must not rename', sender_full_name='Must not change sender', **{field: value})
    before = s.board(USER)
    # DTO rejects coercion/overlong values; format validation remains in the
    # server-authoritative mutation so direct internal callers are safe too.
    if not isinstance(value, str) or len(value) > (4 if field == 'card_last4' else 1024):
        with pytest.raises(ValidationError):
            Client.model_validate(patch)
    with pytest.raises(HTTPException) as exc:
        change('client_save', patch)
    assert exc.value.status_code == 400 and s.board(USER) == before
    assert s.reveal_client_password(USER, client['id']) == 'Original synthetic credential'


def test_missing_encryption_key_rolls_back_new_client_card_events_and_revision(board_db, monkeypatch):
    monkeypatch.delenv('GMAIL_SETUP_PASSWORD_KEY', raising=False)
    monkeypatch.delenv('SESSION_SECRET', raising=False)
    before = s.board(USER)
    with pytest.raises(HTTPException) as exc:
        change('client_save', client_data(password='Must not persist synthetic secret', card_last4='0042', sender_full_name='Not saved'))
    assert exc.value.status_code == 503 and s.board(USER) == before and board_db == {}
    assert not before['state']['clients'] and not before['state']['events'] and before['revision'] == 0


def test_missing_key_and_other_invalid_metadata_each_roll_back_the_whole_edit(board_db, password_key, monkeypatch):
    record = change('client_save', client_data(password='Original synthetic secret', card_last4='0101'))
    client = saved_client(record)
    before = s.board(USER)
    with pytest.raises(HTTPException):
        change('client_save', client_data(id=client['id'], name='Never renamed', password='Never saved', card_last4='0202', website_url='javascript:bad'))
    assert s.board(USER) == before
    monkeypatch.delenv('GMAIL_SETUP_PASSWORD_KEY', raising=False)
    monkeypatch.delenv('SESSION_SECRET', raising=False)
    with pytest.raises(HTTPException) as exc:
        change('client_save', client_data(id=client['id'], name='Still not renamed', password='New synthetic secret', card_last4='0303'))
    assert exc.value.status_code == 503 and s.board(USER) == before


def test_credentials_only_edit_preserves_stats_and_setup_device_dns_sheet_history(board_db, password_key):
    owner = people()['Dylan']
    phone = devices(owner)['state']['devices'][0]
    record, job_id = create_setup(owner, phone['id'])
    client = saved_client(record)
    change('setup_stage', {'id': job_id, 'stage': 'Setup complete'})
    record = s.board(USER)
    record['state']['setups'][0]['dns_setup'] = {'status': 'verified', 'provider_record_id': 'synthetic-record'}
    record['state']['setups'][0]['sync'] = {'status': 'synced', 'row': 42, 'url': 'https://example.com/synthetic-sheet-row'}
    store.save(USER, record['state'], record['revision'])
    before = s.board(USER)
    summary = s.summary(copy.deepcopy(before['state']))
    edited = change('client_save', client_data(id=client['id'], name=client['name'], password='Optional synthetic secret', card_last4='0007'))
    assert edited['state']['setups'] == before['state']['setups'] and edited['state']['devices'] == before['state']['devices']
    assert s.summary(edited['state']) == summary
    assert edited['state']['clients'][0]['next_order'] == before['state']['clients'][0]['next_order']
    assert edited['state']['clients'][0]['domain_target'] == before['state']['clients'][0]['domain_target']


def test_optional_credentials_api_requires_tool_csrf_revision_and_workspace_access(client, board_db, password_key, monkeypatch):
    signed_in = {'user': None}
    monkeypatch.setattr(auth, 'get_session_user', lambda request: signed_in['user'])
    payload = {'revision': 0, 'action': 'client_save', 'data': client_data(password='API synthetic secret', card_last4='0003')}
    endpoint = '/api/gmail-setup/change'
    assert client.post(endpoint, json=payload).status_code == 401
    signed_in['user'] = {**OTHER, 'tools': []}
    assert client.post(endpoint, json=payload).status_code == 403
    signed_in['user'] = {**OTHER}
    assert client.post(endpoint, json=payload).status_code == 403 and board_db == {}
    client.headers['X-Domain-CSRF'] = client.get('/api/gmail-setup/session').json()['csrf']
    response = client.post(endpoint, json=payload)
    assert response.status_code == 200
    public = response.json()['state']['clients'][0]
    assert public['has_password'] is True and public['card_last4'] == '0003' and 'password_ciphertext' not in public
    assert 'API synthetic secret' not in response.text
    identifier = public['id']
    assert client.post(f'/api/gmail-setup/clients/{identifier}/password/reveal').json() == {'password': 'API synthetic secret'}
    before = s.board(USER)
    edit = {'revision': 0, 'action': 'client_save', 'data': client_data(id=identifier, password='Rejected stale secret', card_last4='9999')}
    assert client.post(endpoint, json=edit).status_code == 409 and s.board(USER) == before
    for patch in ({'card_last4': '4242424242424242'}, {'card_last4': '１２３４'}, {'card_last4': 1234}, {'password': True}):
        rejected = client.post(endpoint, json={**edit, 'revision': before['revision'], 'data': {**edit['data'], **patch}})
        assert rejected.status_code == 400 and s.board(USER) == before
        assert 'API synthetic secret' not in rejected.text
    signed_in['user'] = {'email': 'viewer@another-workspace.test', 'role': 'viewer', 'tools': ['gmail_set']}
    foreign = client.post(endpoint, json={**edit, 'revision': 0})
    assert foreign.status_code == 404 and s.board(USER) == before
    assert client.post(f'/api/gmail-setup/clients/{identifier}/password/reveal').status_code == 404
    assert len(board_db) == 1
