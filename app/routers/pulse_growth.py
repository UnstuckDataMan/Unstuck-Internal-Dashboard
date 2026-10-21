"""Client growth planner: staff activation, narrowly token-scoped client editing."""
import json
import logging
import secrets
import os
import hmac
import hashlib
from datetime import date, datetime, timezone
from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import JSONResponse
from app import auth
from app.utils.dates import today_utc
from app.utils import google_sheets
from app.utils.pulse import growth, growth_store, store, tracker
from app.utils.pulse.store import PulseNotReady
from app.routers.pulse_portal import _resolve_access, _private

router = APIRouter()
logger = logging.getLogger(__name__)


def _client(client_id):
    if not any(str(c['id']) == client_id for c in store.list_clients()):
        raise HTTPException(404, 'Client not found.')


def _staff(request, client_id):
    user = auth.require_login(request)
    if auth.auth_enabled() and 'outbound_pulse' not in user.get('tools', []):
        raise HTTPException(403, 'Client reporting access required.')
    _client(client_id)
    return user


# Keep CSRF independent of the signed session cookie. Concurrent dashboard
# requests can return older session cookies and discard newly added fields.
_CSRF_SECRET = os.environ.get('SESSION_SECRET') or secrets.token_urlsafe(32)


def _csrf(request, check=False):
    identity = str((request.session.get('user') or {}).get('email', ''))
    scope = json.dumps([request.url.path, identity], separators=(',', ':'))
    token = hmac.new(_CSRF_SECRET.encode(), scope.encode(), hashlib.sha256).hexdigest()
    if check and not secrets.compare_digest(token, request.headers.get('X-Growth-CSRF', '')):
        raise HTTPException(403, 'Could not verify this save. Please try again; your entries are still here.')
    return token


def _view(request, row, *, staff=False, can_activate=False):
    row = row or {}
    return _private(JSONResponse({'data': row.get('data') or growth.defaults(today_utc()),
        'version': row.get('version', 0), 'enabled': row.get('enabled', False),
        'include_in_reports': row.get('include_in_reports', False),
        'staff': staff, 'can_activate': can_activate, 'csrf': _csrf(request),
        'today': today_utc().isoformat()}))


def _lead_snapshot(client_id, start):
    linked = store.get_tracker(client_id)
    if not linked:
        raise ValueError('Link this client’s Performance Tracker in Lead quality, then refresh leads here.')
    values = google_sheets.read_tracker_values(str(linked.get('sheet_id') or ''),
                                             str(linked.get('tab_title') or tracker.DEFAULT_TAB))
    rows, meta = tracker.parse_rows(values)
    if not meta.get('has_dates'):
        raise ValueError('The lead tracker needs a date column before a forecast can be calculated.')
    end = today_utc()
    selected = tracker.in_period(rows, date.fromisoformat(start), end)
    return {'count': len(selected), 'start': start, 'as_of': end.isoformat(),
            'source': 'Performance Tracker', 'refreshed_at': datetime.now(timezone.utc).isoformat()}


async def _write(request, client_id, *, staff, user=None):
    _csrf(request, check=True)
    raw = await request.body()
    if len(raw) > 200_000:
        raise HTTPException(413, 'This planner is too large.')
    try:
        payload = json.loads(raw)
        if not isinstance(payload, dict):
            raise ValueError('Invalid planner request.')
        previous = growth_store.get(client_id)
        if payload.get('version') != (previous or {}).get('version', 0):
            raise growth_store.Conflict('Someone else saved this planner. Reload before making further changes.')
        action = payload.get('action', 'save')
        if action not in ('save', 'activate', 'deactivate', 'refresh'):
            raise ValueError('Unknown planner action.')
        if not staff and (not previous or not previous.get('enabled') or not previous.get('include_in_reports')):
            raise HTTPException(404, 'Planner unavailable.')
        if not staff and action != 'save':
            raise HTTPException(403, 'Only your account manager can change activation or refresh leads.')
        if action in ('activate', 'deactivate') and (user or {}).get('role') not in ('admin', 'reviewer'):
            raise HTTPException(403, 'An account manager or administrator must activate this planner.')
        data = growth.validate(payload.get('data'), today_utc())
        old_data = (previous or {}).get('data') or {}
        if not staff and data['measurement_start'] != old_data.get('measurement_start'):
            raise HTTPException(403, 'Your account manager sets the lead tracking period.')
        # Lead volumes are server-owned; client-supplied counts are never accepted.
        if data['measurement_start'] == old_data.get('measurement_start') and old_data.get('leads'):
            data['leads'] = old_data['leads']
        if action in ('refresh', 'activate'):
            data['leads'] = _lead_snapshot(client_id, data['measurement_start'])
        enabled = bool((previous or {}).get('enabled'))
        included = bool((previous or {}).get('include_in_reports'))
        approved_by = (previous or {}).get('approved_by', '')
        if action == 'activate':
            if payload.get('delivery_approved') is not True or payload.get('has_conversion') is not True:
                raise ValueError('Confirm delivery manager approval and that the client has had a conversion.')
            if not data['conversions']:
                raise ValueError('Record at least one landed conversion before activating.')
            if not data['leads']['count'] or len(data['conversions']) > data['leads']['count']:
                raise ValueError('Check the tracking period: generated leads must cover the recorded conversions.')
            enabled = True
            included = True
            approved_by = (user or {}).get('email', '')
        elif action == 'deactivate':
            enabled, included = False, False
        if enabled and not data['conversions']:
            raise ValueError('An active planner must have at least one landed conversion.')
        if staff and action == 'save':
            requested = payload.get('include_in_reports', included)
            if not isinstance(requested, bool):
                raise ValueError('Choose whether to include this planner in reports.')
            included = requested and enabled
        saved = growth_store.save(client_id, previous, data, enabled=enabled,
                                  included=included, approved_by=approved_by)
        return _view(request, saved, staff=staff,
                     can_activate=staff and (user or {}).get('role') in ('admin', 'reviewer'))
    except growth_store.Conflict as exc:
        return _private(JSONResponse({'error': str(exc)}, status_code=409))
    except (ValueError, json.JSONDecodeError) as exc:
        return _private(JSONResponse({'error': str(exc)}, status_code=422))
    except HTTPException:
        raise
    except Exception:
        logger.exception('Growth planner save failed for client %s', client_id)
        return _private(JSONResponse({'error': 'Could not save the planner or read the lead tracker. Your changes have not been saved. Please try again.'}, status_code=503))


@router.get('/api/outbound-pulse/clients/{client_id}/growth-plan')
async def staff_get(request: Request, client_id: str):
    user = _staff(request, client_id)
    try:
        return _view(request, growth_store.get(client_id), staff=True,
                     can_activate=user.get('role') in ('admin', 'reviewer'))
    except PulseNotReady as exc:
        return _private(JSONResponse({'error': str(exc)}, status_code=503))


@router.post('/api/outbound-pulse/clients/{client_id}/growth-plan')
async def staff_save(request: Request, client_id: str):
    user = _staff(request, client_id)
    return await _write(request, client_id, staff=True, user=user)


def _portal_client(token, report_id):
    access, _ = _resolve_access(token)
    if not access:
        raise HTTPException(404, 'Planner unavailable.')
    client_id = str(access['client_id'])
    _client(client_id)
    if not any(str(r['id']) == report_id
               for r in store.published_reports(client_id)):
        raise HTTPException(404, 'Planner unavailable.')
    row = growth_store.get(client_id)
    if not row or not row.get('enabled') or not row.get('include_in_reports'):
        raise HTTPException(404, 'Planner unavailable.')
    return client_id, row


@router.get('/r/{token}/growth-plan/{report_id}')
@router.get('/portal/{token}/growth-plan/{report_id}')
async def portal_get(request: Request, token: str, report_id: str):
    try:
        _, row = _portal_client(token, report_id)
        return _view(request, row)
    except HTTPException as exc:
        return _private(JSONResponse({'error': exc.detail}, status_code=exc.status_code))
    except Exception:
        return _private(JSONResponse({'error': 'The growth planner is temporarily unavailable.'}, status_code=503))


@router.post('/r/{token}/growth-plan/{report_id}')
@router.post('/portal/{token}/growth-plan/{report_id}')
async def portal_save(request: Request, token: str, report_id: str):
    try:
        client_id, _ = _portal_client(token, report_id)
        return await _write(request, client_id, staff=False)
    except HTTPException as exc:
        return _private(JSONResponse({'error': exc.detail}, status_code=exc.status_code))
    except Exception:
        return _private(JSONResponse({'error': 'The growth planner is temporarily unavailable.'}, status_code=503))
