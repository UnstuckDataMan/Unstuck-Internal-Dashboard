"""Durable Copy Studio worker: python -m app.utils.copy_studio.worker.

Separate from web processes. A lease and persisted checkpoints allow another
worker to resume after termination. Provider reads are idempotent; no sending.
"""
from __future__ import annotations

import argparse
import logging
import os
import socket
import threading
import time
import uuid
from datetime import datetime, timezone

from . import service
from .store import LeaseLost, StoreError, get_store

logger = logging.getLogger(__name__)


def _persist_items(store, job, result):
    ids = []
    for item in result.get('items', []):
        if item.get('client_id') != job['client_id']:
            raise ValueError('Source returned data for a different client; import refused.')
        record = store.save(item['kind'], job['client_id'], item['source_key'], item['content'], job['created_by'])
        ids.append(record['id'])
    return ids


def _sync(store, job, worker_id):
    from . import sources
    from .integrations import IntegrationUnavailable
    client_id = job['client_id']
    requested = job['payload']['source']
    stages = ['copy_bank', 'onboarding', 'smartlead', 'manual'] if requested == 'all' else [requested]
    checkpoint = job.get('checkpoint') or {}
    completed = list(checkpoint.get('completed', []))
    reports = list(checkpoint.get('reports', []))

    def persist(result):
        # A slow source call must not write after another worker acquired the lease.
        store.heartbeat_job(job['id'], worker_id)
        return _persist_items(store, job, result)

    for source in stages:
        if source in completed:
            continue
        if source in ('copy_bank', 'onboarding'):
            try:
                store.heartbeat_job(job['id'], worker_id)
                result = getattr(sources, 'sync_'+source)(client_id)
                ids = persist(result)
                summary = {'source': source, 'coverage': result.get('coverage', {}),
                           'warnings': result.get('warnings', []), 'records': len(ids)}
            except IntegrationUnavailable as exc:
                summary = {'source': source, 'records': 0,
                           'coverage': {'complete': False}, 'warnings': [str(exc)]}
        else:
            state = checkpoint.get('campaigns', {}) if checkpoint.get('source') == source else {}
            offset, index = state.get('offset', 0), state.get('index', 0)
            cursor, count = state.get('cursor'), state.get('count', 0)
            warnings, coverages = list(state.get('warnings', [])), list(state.get('coverage', []))
            # Persist the enumeration page as well as the per-campaign cursor:
            # a source campaign inserted/deleted during a restart must not move
            # the saved array index onto another campaign's statistics cursor.
            listing = state.get('listing')
            while True:
                store.heartbeat_job(job['id'], worker_id)
                if listing is None:
                    try:
                        listing = sources.list_campaigns(client_id, source, offset=offset, limit=50)
                    except IntegrationUnavailable as exc:
                        warnings.append(str(exc))
                        coverages.append({'complete': False, 'stage': 'campaign_listing', 'error': str(exc)})
                        break
                campaigns = listing.get('campaigns', [])
                for position, campaign in enumerate(campaigns):
                    if position < index:
                        continue
                    while True:
                        store.heartbeat_job(job['id'], worker_id)
                        try:
                            result = sources.sync_campaign(client_id, source, campaign, cursor=cursor, max_pages=5)
                            count += len(persist(result))
                            warnings.extend(result.get('warnings', []))
                            next_cursor = result.get('next_cursor')
                            finished = result.get('complete', True)
                            if not finished and (next_cursor is None or next_cursor == cursor):
                                warnings.append('Campaign history could not advance; remaining history is unresolved.')
                                coverages.append({'campaign_id': str(campaign.get('id')),
                                                  **result.get('coverage', {}), 'complete': False,
                                                  'error': 'Missing or non-advancing continuation cursor.'})
                                finished = True
                            elif finished:
                                coverages.append({'campaign_id': str(campaign.get('id')), **result.get('coverage', {})})
                            cursor = next_cursor
                        except IntegrationUnavailable as exc:
                            # One missing share or deleted provider campaign does
                            # not hide valid evidence from every other campaign.
                            warnings.append(str(exc))
                            coverages.append({'campaign_id': str(campaign.get('id')), 'complete': False, 'error': str(exc)})
                            finished, cursor = True, None
                        checkpoint = {'source': source, 'completed': completed, 'reports': reports,
                                      'campaigns': {'offset': offset, 'index': position+1 if finished else position,
                                                    'cursor': None if finished else cursor, 'count': count,
                                                    'listing': listing,
                                                    'warnings': list(dict.fromkeys(warnings)), 'coverage': coverages}}
                        store.checkpoint_job(job['id'], worker_id, checkpoint)
                        if finished:
                            cursor = None
                            break
                    index = position+1
                if listing.get('complete', True):
                    break
                next_offset = listing.get('next_offset')
                if next_offset is None or next_offset <= offset:
                    warnings.append('Campaign pagination did not advance; listing coverage is partial.')
                    coverages.append({'complete': False, 'stage': 'campaign_listing', 'error': warnings[-1]})
                    break
                offset, index, listing = next_offset, 0, None
                checkpoint = {'source': source, 'completed': completed, 'reports': reports,
                              'campaigns': {'offset': offset, 'index': 0, 'cursor': None, 'count': count,
                                            'listing': None, 'warnings': warnings, 'coverage': coverages}}
                store.checkpoint_job(job['id'], worker_id, checkpoint)
            summary = {'source': source, 'records': count,
                       'coverage': {'campaigns': coverages, 'complete': all(c.get('complete', True) for c in coverages)},
                       'warnings': list(dict.fromkeys(warnings))}
            if source == 'smartlead' and summary['coverage']['complete']:
                try:
                    store.heartbeat_job(job['id'], worker_id)
                    reconciliation = sources.reconcile_smartlead(client_id)
                    summary['reconciliation'] = reconciliation
                    if reconciliation['client_mappings_missing_from_provider']:
                        summary['coverage']['complete'] = False
                        summary['warnings'].append('Some mapped client campaigns are absent from the current provider list; historical records are retained.')
                    if reconciliation['provider_without_agency_mapping_count']:
                        summary['warnings'].append('Some provider campaigns lack an agency client mapping and have not been imported into this client.')
                except IntegrationUnavailable as exc:
                    summary['coverage']['complete'] = False
                    summary['warnings'].append('Campaign reconciliation is incomplete. '+str(exc))
        summary['completed_at'] = datetime.now(timezone.utc).isoformat()
        summary['status'] = 'complete' if summary.get('coverage', {}).get('complete', True) else 'partial'
        # A completed attempt is not automatically a successful source sync.
        if summary['status'] == 'complete':
            summary['last_successful_sync_at'] = summary['completed_at']
        store.heartbeat_job(job['id'], worker_id)
        store.save('sync_report', client_id, 'sync:'+source, summary, job['created_by'])
        reports.append(summary)
        completed.append(source)
        checkpoint = {'completed': completed, 'reports': reports}
        store.checkpoint_job(job['id'], worker_id, checkpoint)
    return {'sources': reports, 'complete': all(r.get('coverage', {}).get('complete', True) for r in reports)}


def execute(store, job, worker_id):
    client_id = job['client_id']
    if job['kind'] == 'sync':
        return _sync(store, job, worker_id)
    if job['kind'] == 'extract_brief':
        from .briefs import extract_brief
        sources = [store.get(i, client_id) for i in job['payload']['source_ids']]
        if any(not s or s['kind'] != 'transcript' for s in sources):
            raise ValueError('An onboarding source is missing or belongs to another client.')
        previous = service.latest_approved(store.iter_records(client_id, kind='brief', status='approved'), 'brief')
        def progress(value):
            store.checkpoint_job(job['id'], worker_id, value)
        content = extract_brief(client_id, sources, previous=previous,
                                progress=progress, checkpoint=job.get('checkpoint'))
        record = store.save('brief', client_id, 'onboarding-brief', content, job['created_by'],
                            parent_id=previous['id'] if previous else None)
        if content.get('errors') and not content.get('coverage', {}).get('stored_text_complete'):
            from .providers import ProviderError
            raise ProviderError('Brief extraction is incomplete. Completed sections are saved; the job will retry from its checkpoint.')
        return {'record_id': record['id'], 'coverage': content.get('coverage'), 'status': 'needs_review'}
    if job['kind'] == 'generate':
        from .writer import generate
        payload = job['payload']
        brief = store.get(payload['brief_id'], client_id)
        examples = [store.get(i, client_id) for i in payload['example_ids']]
        if not brief or brief['status'] != 'approved' or any(not e or e['status'] != 'approved' for e in examples):
            raise ValueError('The brief or references are no longer approved. Review them before retrying.')
        current = service.latest_approved(store.iter_records(client_id, kind='brief', status='approved'), 'brief')
        if current and current['id'] != brief['id']:
            raise ValueError('A newer client brief is approved. Generate again using the current brief.')
        methodology = store.get(payload['methodology_id'], service.GLOBAL_CLIENT) if payload.get('methodology_id') else None
        if methodology and methodology['status'] != 'approved':
            raise ValueError('The selected methodology is no longer approved.')
        content = generate(client_id, brief, examples, payload['request'], methodology=methodology,
                           performance=list(store.iter_records(client_id, kind='performance')))
        content['studio_context'] = payload
        record = store.save('draft', client_id, 'generation:'+job['id'], content, job['created_by'])
        return {'record_id': record['id'], 'status': 'needs_review', 'review': content.get('review', {})}
    if job['kind'] == 'review_draft':
        from .writer import review_draft
        draft = store.get(job['payload']['record_id'], client_id)
        if not draft or draft['kind'] != 'draft':
            raise ValueError('Draft version not found.')
        brief, examples, request = service.draft_context(client_id, draft)
        methodology_id = draft['content'].get('studio_context', {}).get('methodology_id')
        methodology = store.get(methodology_id, service.GLOBAL_CLIENT) if methodology_id else None
        content = review_draft(client_id, brief, examples, request, draft['content'], methodology=methodology,
                               performance=list(store.iter_records(client_id, kind='performance')))
        record = store.save('draft', client_id, draft['source_key'], content, job['created_by'], parent_id=draft['id'])
        return {'record_id': record['id'], 'status': 'needs_review', 'review': content['review']}
    raise ValueError('Unknown Copy Studio job type.')


def run_once(store=None, worker_id=None):
    store = store or get_store()
    worker_id = worker_id or 'studio-'+uuid.uuid4().hex
    job = store.claim_job(worker_id, lease_seconds=300)
    if not job:
        return False
    stopped = threading.Event()
    lease_lost = threading.Event()
    def renew():
        while not stopped.wait(45):
            try:
                store.heartbeat_job(job['id'], worker_id)
            except StoreError:
                lease_lost.set()
                return
    heartbeat = threading.Thread(target=renew, daemon=True)
    heartbeat.start()
    try:
        from . import costs
        with costs.metering(store, job):
            result = execute(store, job, worker_id)
        if lease_lost.is_set():
            raise LeaseLost('The job lease expired before completion.')
        store.complete_job(job['id'], worker_id, result)
    except LeaseLost:
        logger.warning('Copy Studio worker lease lost; the persisted job can be resumed.')
    except Exception as exc:
        # Do not retain upstream bodies/URLs/tokens or prospect text in job errors.
        from .providers import ProviderError
        from .costs import CostRecordingError
        safe = str(exc) if isinstance(exc, (ValueError, StoreError, ProviderError)) else 'Source processing failed. Check provider configuration and retry.'
        try:
            store.fail_job(job['id'], worker_id, safe[:500], retry=not isinstance(exc, (ValueError, CostRecordingError)))
        except LeaseLost:
            logger.warning('Job failure occurred after the lease expired; resume will retry.')
    finally:
        stopped.set()
        heartbeat.join(timeout=1)
    return True


def schedule(store=None, at=None):
    """Hourly refreshes for clients explicitly enrolled by a first source sync."""
    store = store or get_store()
    at = at or datetime.now(timezone.utc)
    count = 0
    for client in service.clients(store):
        subscriptions = store.list(client['id'], kind='sync_subscription', limit=1)
        if not subscriptions or not subscriptions[0]['content'].get('enabled'):
            continue
        for source in ('copy_bank', 'onboarding', 'smartlead', 'manual'):
            store.enqueue(client['id'], 'sync', {'source': source}, 'system:copy-studio-worker',
                          f'scheduled:{source}:{at.strftime("%Y-%m-%dT%H")}')
            count += 1
    return count


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--once', action='store_true', help='Process at most one persisted job.')
    parser.add_argument('--schedule', action='store_true', help='Enrol periodic hourly refreshes for subscribed clients.')
    parser.add_argument('--poll-seconds', type=float, default=3)
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO)
    store = get_store()
    worker_id = socket.gethostname()+':'+uuid.uuid4().hex
    scheduled_hour = None
    while True:
        hour = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H')
        if args.schedule and hour != scheduled_hour:
            schedule(store)
            scheduled_hour = hour
        worked = run_once(store, worker_id)
        if args.once:
            break
        if not worked:
            time.sleep(max(0.5, min(args.poll_seconds, 60)))


if __name__ == '__main__':
    main()
