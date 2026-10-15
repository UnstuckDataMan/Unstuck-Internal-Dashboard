"""Read-only readiness; never submits dummy account rows to the live sheet."""
import json
import logging
import threading
import requests
from app.utils import supabase as sb
from . import sheets


def check():
    ready = False
    if sb.sb_service_configured():
        try:
            r = requests.get(sb.SUPABASE_URL+'/rest/v1/gmail_setup_boards', params={'select': 'workspace_id', 'limit': '0'}, headers=sb.sb_service_headers(), timeout=(3, 8))
            ready = r.ok
        except requests.RequestException:
            pass
    return {'storage_ready': ready, 'sheet': sheets.readiness()}


def start_check():
    def run():
        try:
            logging.getLogger('uvicorn.error').info('Gmail Set tracker check: %s', json.dumps(check(), sort_keys=True))
        except Exception:
            logging.getLogger('uvicorn.error').warning('Gmail Set tracker check unavailable; check sheet access in the tool.')
    threading.Thread(target=run, name='gmail-tracker-check', daemon=True).start()
