"""Real local store + HTTP + worker tests. All source/model responses are fixtures."""
from copy import deepcopy
import json

import pytest

from app import auth
from app.utils.copy_studio import briefs, service, worker, writer
from app.utils.copy_studio.store import get_store

CLIENT = 'client-a'
BASE = '/api/copy-studio/clients/'+CLIENT


@pytest.fixture
def studio(monkeypatch, tmp_path):
    for key in ('COPY_STUDIO_OPENAI_API_KEY', 'COPY_STUDIO_ANTHROPIC_API_KEY',
                'OPENAI_API_KEY', 'ANTHROPIC_API_KEY'):
        monkeypatch.delenv(key, raising=False)
    monkeypatch.setenv('COPY_STUDIO_SQLITE_PATH', str(tmp_path/'studio.sqlite'))
    monkeypatch.setenv('COPY_STUDIO_AGENCY_ID', 'agency-a')
    db = get_store()
    for key in (CLIENT, 'client-b'):
        db.save('client', '__workspace__', key, {'id': key, 'name': 'Synthetic '+key}, 'fixture')
    return db


def source(db, *, complete=True, client_id=CLIENT):
    return db.save('transcript', client_id, 'read_ai:fixture', {
        'source': 'read_ai', 'title': 'Synthetic onboarding fixture',
        'transcript': [{'text': 'We help UK software finance teams simplify reporting.', 'start_ms': 0}],
        'source_complete': complete,
        'source_completeness_review': {'actor': 'reviewer@fixture.test'} if complete else None}, 'fixture')


def approved_brief(db):
    src = source(db)
    content = {
        'client_id': CLIENT, 'facts': [
            {'id': 'offer', 'category': 'offer', 'key': 'offer',
             'text': 'We help UK software finance teams simplify reporting.',
             'evidence': [{'source_id': src['id'], 'block': 0,
                           'quote': 'We help UK software finance teams simplify reporting.'}]},
            {'id': 'audience', 'category': 'audience', 'key': 'audience', 'text': 'UK software finance teams',
             'evidence': [{'source_id': src['id'], 'block': 0, 'quote': 'UK software finance teams'}]},
        ], 'constraints': ['Use restrained language.'], 'contradictions': [],
        'coverage': {'complete': True, 'sources': [{'source_id': src['id'], 'source_version': src['content_hash'],
                                                  'stored_text_complete': True}]} }
    brief = db.save('brief', CLIENT, 'onboarding-brief', content, 'fixture')
    return db.review(brief['id'], CLIENT, 'reviewer@fixture.test', 'approved')


def approved_example(db, channel='email'):
    content = {'source': 'copy_bank', 'channel': channel, 'sanitized': True, 'reusable': True,
               'subjects': ['Simpler reporting'] if channel == 'email' else [],
               'steps': [{'kind': 'initial', 'body': 'Could simpler reporting help your finance team?', 'evidence_ids': ['offer']}],
               'evidence_ids': ['offer']}
    ex = db.save('example', CLIENT, 'copy_bank:fixture:'+channel, content, 'fixture')
    return db.review(ex['id'], CLIENT, 'reviewer@fixture.test', 'approved')


def generation(example, mode='assemble', channel='email'):
    return {'channel': channel, 'mode': mode, 'angles': 1,
            'audience': {'subsector': 'software', 'geography': 'UK', 'role': 'Finance director', 'company_size': '50–200'},
            'example_ids': [example['id']]}


def request_as(monkeypatch, role='sdr', **extra):
    monkeypatch.delenv('AUTH_DISABLED', raising=False)
    user = {'email': role+'@fixture.test', 'role': role, 'tools': ['copy_studio'], **extra}
    monkeypatch.setattr(auth, 'get_session_user', lambda request: user)
    return user


def csrf(client):
    return {'X-Copy-Studio-CSRF': client.get('/api/copy-studio/bootstrap').json()['csrf']}


def test_keyless_assembly_roundtrip_queue_worker_revision_approval(client, studio):
    approved_brief(studio)
    example = approved_example(studio)
    response = client.post(BASE+'/generate', json=generation(example), headers={'Idempotency-Key': 'one'})
    assert response.status_code == 200, response.text
    job = response.json()
    assert job['status'] == 'queued'
    duplicate = client.post(BASE+'/generate', json=generation(example), headers={'Idempotency-Key': 'one'})
    assert duplicate.json()['id'] == job['id']
    assert worker.run_once(studio, 'fixture-worker')
    finished = client.get(BASE+'/jobs/'+job['id']).json()
    assert finished['status'] == 'succeeded', finished
    record_id = finished['result']['record_id']
    draft = client.get(BASE+'/records/'+record_id).json()
    assert draft['content']['provenance']['provider'] == 'none'
    assert draft['status'] == 'pending'
    response = client.post(BASE+'/records/'+record_id+'/review', json={'decision': 'approved'})
    assert response.status_code == 200, response.text
    assert response.json()['status'] == 'approved'
    edited = deepcopy(draft['content'])
    edited['angles'][0]['steps'][0]['body'] = 'Would simpler reporting be useful for your team?'
    revision = client.post(BASE+'/drafts/'+record_id+'/revisions', json={'content': edited})
    assert revision.status_code == 200, revision.text
    assert revision.json()['id'] != record_id
    assert revision.json()['status'] == 'pending'
    assert studio.get(record_id, CLIENT)['content']['angles'][0]['steps'][0]['body'] == 'Could simpler reporting help your finance team?'


def test_client_and_agency_scoping_including_methodology(client, studio, monkeypatch):
    request_as(monkeypatch, 'reviewer', client_ids=[CLIENT])
    token = csrf(client)
    assert len(client.get('/api/copy-studio/bootstrap').json()['clients']) == 1
    assert client.get('/api/copy-studio/clients/client-b').status_code == 404
    other = studio.save('example', 'client-b', 'other', {'body': 'Other client private text'}, 'fixture')
    assert client.get(BASE+'/records/'+other['id']).status_code == 404
    assert client.post(BASE+'/records/'+other['id']+'/review', json={'decision': 'approved'}, headers=token).status_code == 404
    request_as(monkeypatch, 'admin', agency_id='different-agency')
    assert client.post('/api/copy-studio/methodology/seed', headers=token).status_code == 403


def test_csrf_roles_and_approval_actor_cannot_be_spoofed(client, studio, monkeypatch):
    brief = approved_brief(studio)
    request_as(monkeypatch)
    assert client.post(BASE+'/sync', json={}).status_code == 403
    token = csrf(client)
    assert client.post(BASE+'/records/'+brief['id']+'/review', json={'decision': 'approved'}, headers=token).status_code == 403
    request_as(monkeypatch, 'reviewer')
    token = csrf(client)
    spoofed = client.post(BASE+'/records/'+brief['id']+'/review', headers=token,
                          json={'decision': 'approved', 'approved_by': 'ceo@fake.test'})
    assert spoofed.status_code == 422
    response = client.post(BASE+'/records/'+brief['id']+'/review', json={'decision': 'approved'}, headers=token)
    assert response.status_code == 200, response.text
    assert 'reviewer@fixture.test' in json.dumps(response.json())
    request_as(monkeypatch, 'client', client_ids=[CLIENT])
    assert client.get(BASE).status_code == 403


def test_forged_coverage_rejected_and_explicit_review_creates_source_revision(client, studio):
    src = source(studio, complete=False)
    content = {'facts': [{'id': 'x', 'category': 'offer', 'text': 'Simplify reporting',
                         'evidence': [{'source_id': src['id'], 'block': 0, 'quote': 'simplify reporting'}]}],
               'constraints': [], 'coverage': {'complete': True, 'sources': [
                   {'source_id': src['id'], 'source_version': src['content_hash'], 'stored_text_complete': True}]}}
    brief = studio.save('brief', CLIENT, 'onboarding-brief', content, 'fixture')
    response = client.post(BASE+'/records/'+brief['id']+'/review', json={'decision': 'approved'})
    assert response.status_code == 422
    assert 'full-call' in response.json()['error']
    confirmed = client.post(BASE+'/transcripts/'+src['id']+'/confirm-complete',
                            json={'confirmed': True, 'notes': 'Compared against the original full call.'})
    assert confirmed.status_code == 200
    assert confirmed.json()['id'] != src['id']
    assert confirmed.json()['content']['source_completeness_review']['actor'] == 'dev@local'
    assert studio.get(src['id'], CLIENT)['content']['source_complete'] is False


def test_duplicate_fact_ids_and_resolution_notes_do_not_hide_conflicts(client, studio):
    base = approved_brief(studio)
    content = deepcopy(base['content'])
    content['facts'][1]['id'] = content['facts'][0]['id']
    invalid = studio.save('brief', CLIENT, 'invalid', content, 'fixture')
    response = client.post(BASE+'/records/'+invalid['id']+'/review', json={'decision': 'approved'})
    assert response.status_code == 422
    assert 'unique' in response.json()['error']
    content = deepcopy(base['content'])
    content['facts'].append(dict(content['facts'][1], id='conflict', text='Target every sector.'))
    content['contradictions'] = [{'key': 'audience:audience', 'resolution': 'Please use the latest.'}]
    invalid = studio.save('brief', CLIENT, 'contradictory', content, 'fixture')
    response = client.post(BASE+'/records/'+invalid['id']+'/review', json={'decision': 'approved'})
    assert response.status_code == 422
    assert 'Conflicting facts' in response.json()['error']


def test_brief_edit_cannot_replace_coverage_and_missing_sources_block_generation(client, studio):
    brief = approved_brief(studio)
    content = deepcopy(brief['content'])
    content['coverage'] = {'complete': False, 'sources': []}
    edited = client.post(BASE+'/records', json={'kind': 'brief', 'source_key': 'onboarding-brief',
                                               'parent_id': brief['id'], 'content': content})
    assert edited.status_code == 200
    assert edited.json()['content']['coverage'] == brief['content']['coverage']
    assert client.post('/api/copy-studio/clients/client-b/generate', json=generation({'id': 'none'})).status_code == 422


def test_linkedin_cannot_use_email_or_promoted_smartlead_example(client, studio):
    approved_brief(studio)
    email = approved_example(studio)
    response = client.post(BASE+'/generate', json=generation(email, channel='linkedin'))
    assert response.status_code == 422
    example = approved_example(studio, 'linkedin')
    response = client.post(BASE+'/generate', json=generation(example, channel='linkedin'))
    assert response.status_code == 200
    worker.run_once(studio, 'linkedin-worker')
    assert studio.get_job(response.json()['id'], CLIENT)['status'] == 'succeeded'


def test_sent_copy_requires_sanitised_reference_before_approval(client, studio):
    sent = studio.save('sent_copy', CLIENT, 'smartlead:campaign:1:copy:hash',
                       {'source': 'smartlead', 'campaign_id': '1', 'variant_id': '100',
                        'subject': 'For Bob', 'body': 'Private personalised copy for Bob'}, 'fixture')
    assert client.post(BASE+'/records/'+sent['id']+'/review', json={'decision': 'approved'}).status_code == 422
    response = client.post(BASE+'/records/'+sent['id']+'/reference',
                           json={'subject': 'Reporting', 'body': 'Could simpler reporting help?', 'sanitized': True})
    assert response.status_code == 200
    prepared = response.json()
    assert prepared['status'] == 'pending'
    assert prepared['content']['original_sent_record_id'] == sent['id']
    assert 'Bob' not in prepared['content']['steps'][0]['body']
    assert client.post(BASE+'/records/'+prepared['id']+'/review', json={'decision': 'approved'}).status_code == 200


def test_missing_configuration_is_actionable_no_synthetic_clients(client, monkeypatch):
    monkeypatch.delenv('COPY_STUDIO_SQLITE_PATH', raising=False)
    monkeypatch.delenv('SUPABASE_SERVICE_ROLE_KEY', raising=False)
    payload = client.get('/api/copy-studio/bootstrap').json()
    assert payload['ready'] is False and payload['clients'] == []
    assert payload['sending_enabled'] is False
    assert 'service' in payload['error'].lower()
    alfred = next(c for c in payload['capabilities'] if c['id'] == 'meet_alfred')
    assert alfred['status'] == 'disabled' and not alfred['configured']


def test_model_missing_leaves_persisted_failed_job_not_fabricated_copy(client, studio, monkeypatch):
    approved_brief(studio)
    example = approved_example(studio)
    monkeypatch.delenv('OPENAI_API_KEY', raising=False)
    monkeypatch.delenv('ANTHROPIC_API_KEY', raising=False)
    job = client.post(BASE+'/generate', json=generation(example, mode='draft')).json()
    worker.run_once(studio, 'no-model')
    record = studio.get_job(job['id'], CLIENT)
    assert record['status'] in ('queued', 'failed')
    assert not studio.list(CLIENT, kind='draft')
    assert 'credential' in record.get('error', '') or 'credential' in record.get('last_error', '')


def test_new_draft_edit_is_saved_but_requires_fresh_model_check(client, studio, monkeypatch):
    from app.utils.copy_studio.providers import ProviderError
    approved_brief(studio)
    example = approved_example(studio)
    class Model:
        name, model = 'fixture', 'synthetic'
        def complete_json(self, system, user, **kw):
            if kw['schema_name'] == 'copy_studio_review':
                return {'issues': []}
            return {'angles': [{'name': 'Reporting', 'rationale': 'An operational angle.',
                'subjects': ['Reporting', 'Simpler reporting'], 'steps': [
                    {'kind': 'initial', 'body': 'Could simpler reporting help?', 'evidence_ids': ['offer']},
                    {'kind': 'follow_up', 'body': 'Is reporting a priority for your team?', 'evidence_ids': ['offer']},
                    {'kind': 'closing', 'body': 'Happy to leave this here if the timing is wrong.', 'evidence_ids': []}],
                'evidence_ids': ['offer'], 'claims': []}]}
    monkeypatch.setattr(writer, 'get_provider', lambda: Model())
    job = client.post(BASE+'/generate', json=generation(example, mode='draft')).json()
    worker.run_once(studio, 'model-worker')
    done = studio.get_job(job['id'], CLIENT)
    assert done['status'] == 'succeeded', done
    draft = studio.get(done['result']['record_id'], CLIENT)
    content = deepcopy(draft['content'])
    content['angles'][0]['steps'][0]['body'] = 'Would simpler reporting help your finance team?'
    edited = client.post(BASE+'/drafts/'+draft['id']+'/revisions', json={'content': content})
    assert edited.status_code == 200, edited.text
    version = edited.json()
    assert version['status'] == 'pending'
    assert any(i['code'] == 'model_review_stale' for i in version['content']['review']['issues'])
    assert client.post(BASE+'/records/'+version['id']+'/review', json={'decision': 'approved'}).status_code == 422
    check = client.post(BASE+'/drafts/'+version['id']+'/check', json={}).json()
    worker.run_once(studio, 'review-worker')
    checked = studio.get_job(check['id'], CLIENT)
    assert checked['status'] == 'succeeded', checked
    response = client.post(BASE+'/records/'+checked['result']['record_id']+'/review', json={'decision': 'approved'})
    assert response.status_code == 200, response.text


def test_page_and_bootstrap_do_not_expose_credentials(client, studio, monkeypatch):
    monkeypatch.setenv('OPENAI_API_KEY', 'never-render-this-secret')
    html = client.get('/copy-studio')
    assert html.status_code == 200
    assert 'AI Copy Studio' in html.text
    assert 'never-render-this-secret' not in html.text+client.get('/api/copy-studio/bootstrap').text
    assert html.headers['cache-control'] == 'no-store'


def test_usage_cards_include_worker_calls_and_scope_to_authorised_clients(client, studio, monkeypatch):
    from app.utils.copy_studio import providers, costs
    approved_brief(studio)
    example = approved_example(studio)
    monkeypatch.setenv('COPY_STUDIO_MODEL', 'fixture-model')
    monkeypatch.setenv('COPY_STUDIO_OPENAI_API_KEY', 'fixture-key')
    monkeypatch.setenv('COPY_STUDIO_PRICING_JSON', json.dumps({'openai:fixture-model': {'input': 2, 'output': 6}}))
    def post(url, **kwargs):
        schema = kwargs['json']['text']['format']['name']
        out = {'issues': []} if schema == 'copy_studio_review' else {'angles': [{
            'name': 'Reporting', 'rationale': 'Simplify the reporting work.', 'subjects': ['Reporting', 'Simpler reporting'],
            'steps': [{'kind': 'initial', 'body': 'Could simpler reporting help?', 'evidence_ids': ['offer']},
                      {'kind': 'follow_up', 'body': 'Is reporting a priority?', 'evidence_ids': ['offer']},
                      {'kind': 'closing', 'body': 'Happy to leave this here for now.', 'evidence_ids': []}],
            'evidence_ids': ['offer'], 'claims': []}]}
        class Response:
            status_code, headers = 200, {}
            def json(self):
                return {'id': schema, 'status': 'completed', 'usage': {'input_tokens': 1000, 'output_tokens': 500},
                        'output': [{'content': [{'type': 'output_text', 'text': json.dumps(out)}]}]}
        return Response()
    monkeypatch.setattr(providers.requests, 'post', post)
    job = service.enqueue_generation(CLIENT, generation(example, mode='draft'), {'email': 'fixture'}, 'priced')
    worker.run_once(studio, 'priced-worker')
    assert studio.get_job(job['id'], CLIENT)['status'] == 'succeeded'
    summary = client.get('/api/copy-studio/costs').json()
    for key in ('today', 'month', 'all_time'):
        assert summary[key]['amount'] == pytest.approx(0.01)
        assert summary[key]['total_calls'] == 2
    with costs.metering(studio, {'id': 'other-job', 'client_id': 'client-b', 'created_by': 'fixture'}):
        costs.record_usage('openai', 'fixture-model', {'input_tokens': 1000000, 'output_tokens': 0})
    request_as(monkeypatch, 'sdr', client_ids=[CLIENT])
    assert client.get('/api/copy-studio/costs').json()['all_time']['amount'] == pytest.approx(0.01)
