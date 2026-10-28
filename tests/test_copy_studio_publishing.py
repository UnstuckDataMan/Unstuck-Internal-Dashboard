from copy import deepcopy

import pytest

from app.utils.copy_studio import publishing, service
from app.utils.copy_studio.proofreading import copy_hash
from app.utils.copy_studio.store import SQLiteStore, StoreNotReady


class Gateway:
    def __init__(self, client='client-a', kind='client'):
        self.profile = {'client_id': client, 'name': 'Fixture client', 'type': kind}
        self.content = {'email': {'subjects': ['Existing subject'], 'variations': [{'body': 'Existing opening'}]},
                        'linkedin': {'steps': [{'body': 'Existing LinkedIn message'}]}, 'private_metadata': {'keep': True}}
        self.pending = False
        self.calls = []

    def destinations(self, client_id):
        return {'ok': True, 'profile': self.profile, 'territories': ['UK'],
                'industries': [{'value': 'Software', 'label': 'Software & IT'}]}

    def destination(self, client_id, selection):
        if selection.get('territory') != 'UK' or selection.get('industry') != 'Software':
            return {'ok': False, 'message': 'Choose from the exact client profile.'}
        return {'ok': True, 'profile': self.profile, 'key': '__c__'+client_id+'__UK_Software',
                'content': deepcopy(self.content), 'version': 'destination-v1', 'pending': self.pending}

    def publish(self, **kwargs):
        self.calls.append(kwargs)
        return {'published': True, 'already_published': False, 'publication_id': kwargs['publication_id']}


@pytest.fixture
def scenario(tmp_path):
    store = SQLiteStore(str(tmp_path/'publish.sqlite'), 'agency-a')
    brief = store.save('brief', 'client-a', 'brief', {'coverage': {'sources': []}}, 'author@example.test')
    store.review(brief['id'], 'client-a', 'reviewer@example.test', 'approved')
    example = store.save('example', 'client-a', 'example-1', {'source': 'copy_bank'}, 'importer@example.test')
    store.review(example['id'], 'client-a', 'reviewer@example.test', 'approved')
    content = {'channel': 'email', 'mode': 'assemble', 'request': {'channel': 'email', 'locale': 'en-GB'},
               'angles': [{'name': 'Clarity', 'subjects': ['Less reporting work', 'Reporting question'],
                           'steps': [{'kind': 'initial', 'body': 'Hi {{first_name}}, could simpler reporting help?'},
                                     {'kind': 'follow_up', 'body': 'Would a short conversation help?'},
                                     {'kind': 'closing', 'body': 'Happy to leave this here.'}]}],
               'review': {'can_approve': True, 'issues': []},
               'studio_context': {'brief_id': brief['id'], 'example_ids': [example['id']],
                                  'request': {'channel': 'email', 'locale': 'en-GB'}}}
    def save_draft(value, source='generation-one', parent=None, approve=True):
        value = deepcopy(value)
        digest = copy_hash(value)
        value['final_check'] = {'complete': True, 'locale': 'en-GB', 'copy_hash': digest,
                                'checked_angles': deepcopy(value['angles']),
                                'report': {'complete': True, 'locale': 'en-GB', 'copy_hash': digest, 'issues': []}}
        draft = store.save('draft', 'client-a', source, value, 'author@example.test', parent_id=parent)
        if approve:
            store.review(draft['id'], 'client-a', 'reviewer@example.test', 'approved')
        return store.get(draft['id'], 'client-a')
    draft = save_draft(content)
    return store, brief, example, draft, save_draft, Gateway()


def selection(**extra):
    return {'territory': 'UK', 'industry': 'Software', 'section': 'email', 'angle_index': 0, **extra}


def test_destinations_are_exact_client_profile_and_flyout_is_bizdev_only(scenario):
    store, _, _, _, _, gateway = scenario
    options = publishing.destinations('client-a', store=store, gateway=gateway)
    assert options['configured'] is True
    assert {s['id'] for s in options['sections']} == {'email', 'linkedin', 'chaser'}
    gateway.profile['type'] = 'bizdev'
    assert 'flyout' in {s['id'] for s in publishing.destinations('client-a', store=store, gateway=gateway)['sections']}
    gateway.profile['client_id'] = 'other-client'
    with pytest.raises(publishing.PublishingError, match='exact profile'):
        publishing.destinations('client-a', store=store, gateway=gateway)


def test_preview_exact_approved_text_canonical_selection_and_idempotency_key(scenario):
    store, _, _, draft, _, gateway = scenario
    first = publishing.preview('client-a', draft['id'], selection(), store=store, gateway=gateway)
    again = publishing.preview('client-a', draft['id'], selection(step_indices=[0]), store=store, gateway=gateway)
    assert first['publication_id'] == again['publication_id']
    assert first['selection']['step_indices'] == [0]
    assert first['appended_copy']['subjects'] == draft['content']['angles'][0]['subjects']
    assert first['appended_copy']['steps'] == [{'kind': 'initial', 'body': 'Hi {{first_name}}, could simpler reporting help?'}]
    assert first['append_summary'] == {'subjects_added': 2, 'bodies_added': 1, 'existing_subjects': 1, 'existing_bodies': 1}
    assert gateway.calls == []
    assert gateway.content['private_metadata'] == {'keep': True}


def test_publish_passes_review_identity_scope_snapshot_and_preview_version_without_browser_copy(scenario):
    store, brief, example, draft, _, gateway = scenario
    out = publishing.publish('client-a', draft['id'], selection(body='browser injection'), 'preview-original',
                             'reviewer@example.test', store=store, gateway=gateway)
    call = gateway.calls[0]
    assert out['published']
    assert call['expected_version'] == 'preview-original'
    assert call['actor'] == 'reviewer@example.test'
    assert 'body' not in call['selection']
    assert set(call['context_snapshot']) == {brief['id'], example['id'], draft['id']}
    assert all(r['status'] == 'approved' and r['review_id'] for r in call['context_snapshot'].values())


@pytest.mark.parametrize('change', [
    {'angle_index': True}, {'angle_index': -1}, {'angle_index': 20},
    {'step_indices': []}, {'step_indices': [True]}, {'step_indices': [0, 0]},
    {'step_indices': [1]}, {'section': 'linkedin'}, {'section': 'flyout'}, {'territory': 'Other'},
])
def test_invalid_or_incompatible_selections_fail_closed(scenario, change):
    store, _, _, draft, _, gateway = scenario
    with pytest.raises(publishing.PublishingError):
        publishing.preview('client-a', draft['id'], selection(**change), store=store, gateway=gateway)
    assert not gateway.calls


def test_followups_append_only_compatible_steps_and_warn_about_existing_sequence(scenario):
    store, _, _, draft, _, gateway = scenario
    gateway.content['chaser'] = [{'body': 'Existing follow-up'}]
    out = publishing.preview('client-a', draft['id'], selection(section='chaser'), store=store, gateway=gateway)
    assert out['selection']['step_indices'] == [1, 2]
    assert out['appended_copy']['subjects'] == []
    assert len(out['appended_copy']['steps']) == 2
    assert out['warnings']


def test_linkedin_routes_only_to_linkedin_without_email_subjects(scenario):
    store, _, _, draft, save, gateway = scenario
    content = deepcopy(draft['content'])
    content['channel'] = content['request']['channel'] = 'linkedin'
    content['angles'][0]['steps'].insert(0, {'kind': 'connection', 'body': 'Good to connect, {{first_name}}.'})
    linkedin = save(content, source='linkedin')
    out = publishing.preview('client-a', linkedin['id'], selection(section='linkedin'), store=store, gateway=gateway)
    assert len(out['appended_copy']['steps']) == 4
    assert out['appended_copy']['subjects'] == []
    assert out['warnings']


def test_new_unapproved_edit_blocks_previously_approved_draft(scenario):
    store, _, _, draft, save, gateway = scenario
    edited = deepcopy(draft['content'])
    edited['angles'][0]['steps'][0]['body'] = 'A later unapproved edit.'
    save(edited, parent=draft['id'], approve=False)
    with pytest.raises(publishing.PublishingError, match='newer edited'):
        publishing.preview('client-a', draft['id'], selection(), store=store, gateway=gateway)


@pytest.mark.parametrize('field', ['copy_hash', 'checked_angles', 'report', 'locale'])
def test_stale_or_incomplete_final_check_cannot_publish(scenario, field):
    store, _, _, draft, _, gateway = scenario
    content = deepcopy(draft['content'])
    content['final_check'].pop(field)
    invalid = store.save('draft', 'client-a', 'invalid-check', content, 'author@example.test')
    store.review(invalid['id'], 'client-a', 'reviewer@example.test', 'approved')
    with pytest.raises(publishing.PublishingError, match='final copy'):
        publishing.preview('client-a', invalid['id'], selection(), store=store, gateway=gateway)


@pytest.mark.parametrize('kind', ['brief', 'example', 'draft'])
def test_revoked_dependency_or_draft_blocks_publish(scenario, kind):
    store, brief, example, draft, _, gateway = scenario
    record = {'brief': brief, 'example': example, 'draft': draft}[kind]
    store.review(record['id'], 'client-a', 'reviewer@example.test', 'rejected')
    with pytest.raises(publishing.PublishingError):
        publishing.preview('client-a', draft['id'], selection(), store=store, gateway=gateway)


def test_new_approved_brief_blocks_old_copy_and_other_client_cannot_publish(scenario):
    store, _, _, draft, _, gateway = scenario
    with pytest.raises(publishing.PublishingError, match='exact profile'):
        publishing.preview('client-b', draft['id'], selection(), store=store, gateway=gateway)
    newer = store.save('brief', 'client-a', 'brief', {'facts': ['changed']}, 'author@example.test')
    store.review(newer['id'], 'client-a', 'reviewer@example.test', 'approved')
    with pytest.raises(publishing.PublishingError, match='brief has changed'):
        publishing.preview('client-a', draft['id'], selection(), store=store, gateway=gateway)


def test_pending_legacy_approval_or_active_context_refresh_blocks_append(scenario):
    store, _, _, draft, _, gateway = scenario
    gateway.pending = True
    with pytest.raises(publishing.PublishingConflict, match='pending approval'):
        publishing.preview('client-a', draft['id'], selection(), store=store, gateway=gateway)
    gateway.pending = False
    store.enqueue('client-a', 'sync', {'source': 'context'}, 'author@example.test', 'sync')
    with pytest.raises(publishing.PublishingError, match='refreshing'):
        publishing.preview('client-a', draft['id'], selection(), store=store, gateway=gateway)


def test_current_corpus_does_not_include_unrelated_cost_or_performance_churn(scenario):
    store, _, _, draft, _, gateway = scenario
    store.save('performance', 'client-a', 'new-result', {'rate': 0.1}, 'worker@example.test')
    store.save('cost', 'client-a', 'usage', {'tokens': 12}, 'worker@example.test')
    publishing.publish('client-a', draft['id'], selection(), 'version', 'reviewer@example.test', store=store, gateway=gateway)
    assert len(gateway.calls[0]['context_snapshot']) == 3


def test_sqlite_never_silently_publishes_to_live_copy_bank(scenario):
    store, _, _, draft, _, _ = scenario
    with pytest.raises(StoreNotReady, match='Local records are not published'):
        publishing.preview('client-a', draft['id'], selection(), store=store)


def test_database_conflict_is_clear_and_does_not_claim_published(scenario):
    store, _, _, draft, _, gateway = scenario
    gateway.publish = lambda **kw: {'ok': False, 'code': 'conflict', 'message': 'Preview again.'}
    with pytest.raises(publishing.PublishingConflict, match='Preview again'):
        publishing.publish('client-a', draft['id'], selection(), 'old-version', 'reviewer@example.test', store=store, gateway=gateway)


def test_malformed_existing_section_fails_without_claiming_a_preview(scenario):
    store, _, _, draft, _, gateway = scenario
    gateway.content['email']['variations'] = None
    with pytest.raises(publishing.PublishingError, match='unsupported format'):
        publishing.preview('client-a', draft['id'], selection(), store=store, gateway=gateway)
