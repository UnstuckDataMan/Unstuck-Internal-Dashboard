"""Read-only tracker storage readiness; strategy sheet updates are manual."""
import json
import logging
import threading
import requests
from app.utils import supabase as sb


def check():
    ready = False
    if sb.sb_service_configured():
        try:
            r = requests.get(sb.SUPABASE_URL+'/rest/v1/gmail_setup_boards', params={'select': 'workspace_id', 'limit': '0'}, headers=sb.sb_service_headers(), timeout=(3, 8))
            ready = r.ok
        except requests.RequestException:
            pass
    return {'storage_ready': ready, 'sheet': {'mode': 'manual', 'message': 'Master Strategy sheet updates are manual.'}}


def start_check():
    def run():
        try:
            logging.getLogger('uvicorn.error').info('Gmail Set tracker check: %s', json.dumps(check(), sort_keys=True))
        except Exception:
            logging.getLogger('uvicorn.error').warning('Gmail Set tracker storage check unavailable.')
    threading.Thread(target=run, name='gmail-tracker-check', daemon=True).start()
