"""Private, agency/client-scoped planner storage with optimistic concurrency."""
from datetime import datetime, timezone
import requests
from app.utils.supabase import SUPABASE_URL, sb_service_headers, sb_service_configured
from app.utils.pulse.store import PulseNotReady, current_agency_id


class Conflict(ValueError):
    pass


def _params(client_id):
    if not sb_service_configured():
        raise PulseNotReady('Growth planner storage is not configured.')
    return {'agency_id': f'eq.{current_agency_id()}', 'client_id': f'eq.{client_id}'}


def get(client_id):
    response = requests.get(f'{SUPABASE_URL}/rest/v1/pulse_growth_plans',
                            headers=sb_service_headers(),
                            params={**_params(client_id), 'select': '*', 'limit': '1'}, timeout=15)
    if response.status_code >= 400:
        raise PulseNotReady('Growth planner storage is unavailable. Check its migration and service key.')
    rows = response.json()
    return rows[0] if rows else None


def save(client_id, previous, data, *, enabled, included, approved_by):
    params = _params(client_id)
    body = {'data': data, 'enabled': enabled, 'include_in_reports': included,
            'approved_by': approved_by, 'version': (previous or {}).get('version', 0) + 1,
            'updated_at': datetime.now(timezone.utc).isoformat()}
    if previous:
        response = requests.patch(f'{SUPABASE_URL}/rest/v1/pulse_growth_plans',
            headers=sb_service_headers('return=representation'),
            params={**params, 'version': f'eq.{previous["version"]}'}, json=body, timeout=15)
    else:
        response = requests.post(f'{SUPABASE_URL}/rest/v1/pulse_growth_plans',
            headers=sb_service_headers('return=representation'),
            json={**body, 'agency_id': current_agency_id(), 'client_id': client_id}, timeout=15)
    if response.status_code == 409:
        raise Conflict('Someone else saved this planner. Reload before making further changes.')
    if response.status_code >= 400:
        raise PulseNotReady('Could not save the growth planner. Your changes have not been saved.')
    rows = response.json()
    if not rows:
        raise Conflict('Someone else saved this planner. Reload before making further changes.')
    return rows[0]
