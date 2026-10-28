"""The legacy approval endpoint trusts reviewer sessions, never submitted names."""
import pytest

from app import auth
from app.routers import copy_bank
from conftest import FakeResponse


@pytest.fixture
def approval_env(monkeypatch, fake_sb):
    monkeypatch.delenv('AUTH_DISABLED', raising=False)
    notifications = []
    monkeypatch.setattr(copy_bank, '_send_slack_approval_notification',
                        lambda *args: notifications.append(args))
    pending = {'key': '__c__client-a__UK_PR', 'client_name': 'Client A',
               'territory': 'UK', 'industry': 'PR', 'content': {'email': 'Stored draft'},
               'created_at': '2026-10-28T10:00:00Z', 'requested_by': 'editor@example.com'}
    fake_sb.route('GET', 'copy_bank_pending', lambda call: FakeResponse(json_data=[pending]))
    return fake_sb, notifications, pending


def login(monkeypatch, role, *, tools=None, email='reviewer@unstuck-agency.com'):
    monkeypatch.setattr(auth, 'get_session_user', lambda request: {
        'email': email, 'name': 'Reviewer', 'role': role,
        'tools': ['copy_bank'] if tools is None else tools,
    })


@pytest.mark.parametrize('role', ['admin', 'reviewer'])
@pytest.mark.parametrize('submitted_identity', ['Someone Else', None])
def test_approval_attributes_session_and_uses_stored_content(client, monkeypatch, approval_env,
                                                          role, submitted_identity):
    fake_sb, notifications, pending = approval_env
    login(monkeypatch, role)
    body = {'key': pending['key'], 'content': {'email': 'Forged submitted content'}}
    if submitted_identity is not None:
        body['approved_by'] = submitted_identity
    response = client.post('/api/copy-bank/approve', json=body)
    assert response.status_code == 200
    assert response.json()['ok'] is True
    publication = fake_sb.calls_to('POST', 'copy_bank_templates')[0]['json']
    audit = fake_sb.calls_to('POST', 'copy_approval_logs')[0]['json']
    assert publication['content'] == pending['content']
    assert audit['content_snapshot'] == pending['content']
    assert audit['approved_by'] == 'reviewer@unstuck-agency.com'
    assert notifications == [('Client A', 'UK', 'PR', 'reviewer@unstuck-agency.com')]


@pytest.mark.parametrize('role', ['sdr', 'viewer', 'client', 'unknown'])
def test_nonreviewers_cannot_approve_by_submitting_admin_name(client, monkeypatch, approval_env, role):
    fake_sb, notifications, pending = approval_env
    login(monkeypatch, role)
    response = client.post('/api/copy-bank/approve', json={
        'key': pending['key'], 'content': {}, 'approved_by': 'Admin',
    })
    assert response.status_code == 403
    assert fake_sb.calls == []
    assert notifications == []


@pytest.mark.parametrize('user_kind,expected_status', [('anonymous', 401), ('no_tool', 403), ('no_email', 403)])
def test_missing_identity_or_tool_prevents_all_approval_effects(client, monkeypatch, approval_env,
                                                            user_kind, expected_status):
    fake_sb, notifications, pending = approval_env
    if user_kind == 'anonymous':
        monkeypatch.setattr(auth, 'get_session_user', lambda request: None)
    else:
        login(monkeypatch, 'reviewer', tools=[] if user_kind == 'no_tool' else ['copy_bank'],
              email='' if user_kind == 'no_email' else 'reviewer@unstuck-agency.com')
    response = client.post('/api/copy-bank/approve', json={'key': pending['key'], 'content': {}})
    assert response.status_code == expected_status
    assert fake_sb.calls == []
    assert notifications == []


@pytest.mark.parametrize('role,visible', [('admin', True), ('reviewer', True), ('sdr', False)])
def test_copy_bank_page_exposes_approval_only_to_reviewers(client, monkeypatch, approval_env, role, visible):
    login(monkeypatch, role)
    response = client.get('/copy-bank')
    assert response.status_code == 200
    assert f"const IS_APPROVER       = {'true' if visible else 'false'};" in response.text
