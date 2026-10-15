"""Read-only integration readiness. Reports booleans/counts, never credentials."""
import json
import logging
import threading

import requests
from fastapi import HTTPException
from app.utils import supabase as sb
from . import engine as e, sheets, provisioning


def check():
    report={'checked_at':e.now()}
    try:
        report['porkbun']=provisioning.Porkbun().check()
    except HTTPException as exc:
        report['porkbun']={'credentials_valid':False,'read':False,'dns':False,'message':str(exc.detail)}
    try:
        if not sb.sb_service_configured():
            raise ValueError()
        r=requests.get(sb.SUPABASE_URL+'/rest/v1/domain_cluster_portfolios',params={'select':'id','limit':'0'},headers=sb.sb_service_headers(),timeout=(3,8))
        r.raise_for_status()
        report['storage']={'ready':True}
    except (ValueError,requests.RequestException):
        report['storage']={'ready':False,'message':'Check the Supabase service key and portfolio migration.'}
    try:
        source=sheets.read_live()
        report['sheets']={'ready':not source['invalid'],'live_domains':len({r['domain'] for r in source['records']}),'invalid_live_rows':len(source['invalid'])}
        if source['invalid']:
            report['sheets']['issues']=source['invalid']
            report['sheets']['message']='Correct Live rows before generation: '+', '.join(r['sheet']+' row '+str(r['row']) for r in source['invalid'])
    except HTTPException as exc:
        report['sheets']={'ready':False,'message':str(exc.detail)}
    return report


def start_check():
    def run():
        try:
            report=check()
            logging.getLogger('uvicorn.error').info('Gmail Set integration check: %s',json.dumps(report,sort_keys=True))
        except Exception:
            logging.getLogger('uvicorn.error').warning('Gmail Set integration check could not complete; use Check setup in the tool.')
    threading.Thread(target=run,name='gmail-set-startup-check',daemon=True).start()
