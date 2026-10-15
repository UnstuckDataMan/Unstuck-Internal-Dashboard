"""Local synthetic QA harness only. Never imported by the production application."""
import copy
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
os.environ['AUTH_DISABLED']='1'
os.environ['SESSION_SECRET']='local-qa-only'
from app.main import app
from app import auth
from app.utils.domain_cluster import store
from fastapi import HTTPException
user={'email':'qa@unstuck-agency.com','name':'QA','tools':['gmail_set'],'role':'sdr'}
auth.get_session_user=lambda request:user
rows={}
def save(user,pid,name,state,revision):
    import uuid
    pid=pid or str(uuid.uuid4())
    row={'id':pid,'name':name,'state':copy.deepcopy(state),'revision':revision+1,'updated_at':store.now()}
    rows[pid]=row
    return copy.deepcopy(row)
def load(user,pid):
    if pid not in rows:raise HTTPException(404,'Not found')
    return copy.deepcopy(rows[pid])
store.save=save
store.load=load
store.list_portfolios=lambda user:[{k:v[k] for k in ('id','name','revision','updated_at')} for v in rows.values()]
from app.utils.domain_cluster import sheets
sheet_reads=0
def read_live():
    global sheet_reads
    sheet_reads+=1
    records=[{'domain':'cedarpartners.com','source_kind':'client','source_owner':'Example client','sheet_at_risk':False,'sheet_needs_deleting':False,'sheet_replacing':False},{'domain':'summitgrowth.com','source_kind':'bizdev','source_owner':'Example team','sheet_at_risk':True,'sheet_needs_deleting':False,'sheet_replacing':False}]
    return {'records':records,'invalid':[],'tabs':[{'kind':'client','live_rows':1,'unique_domains':1,'at_risk_rows':0,'needs_deleting_rows':0},{'kind':'bizdev','live_rows':1,'unique_domains':1,'at_risk_rows':1,'needs_deleting_rows':0}],'checked_at':store.now(),'source_url':sheets.SOURCE_URL}
sheets.read_live=read_live
# Synthetic registrar used only by the browser journey, never production.
from app.utils.domain_cluster import provisioning, diagnostics
class Registrar:
    def __init__(self):self.forwards={};self.records={}
    def check(self):return {'credentials_valid':True,'read':True,'dns':True}
    def call(self,method,path,data=None,idempotency=None):
        domain=path.split('/')[-1]
        if path.startswith('/domain/get/'):return {'domain':{'domain':domain,'status':'ACTIVE','apiAccess':1}}
        if path.startswith('/domain/getNs/'):return {'ns':['curitiba.ns.porkbun.com','maceio.ns.porkbun.com']}
        if path.startswith('/domain/getUrlForwarding/'):return {'forwards':copy.deepcopy(self.forwards.get(domain,[]))}
        if path.startswith('/dns/retrieve/'):return {'records':copy.deepcopy(self.records.get(domain,[]))}
        if path.startswith('/domain/addUrlForward/'):
            self.forwards.setdefault(domain,[]).append({**data,'id':'123'});return {'status':'SUCCESS'}
        if path.startswith('/dns/create/'):
            self.records.setdefault(domain,[]).append({**data,'name':data['name']+'.'+domain,'id':'456'});return {'status':'SUCCESS','id':'456'}
        raise RuntimeError('Unexpected synthetic registrar operation')
registrar=Registrar()
provisioning.Porkbun=lambda:registrar
diagnostics.check=lambda:{'porkbun':registrar.check(),'storage':{'ready':True},'sheets':{'ready':True}}
user['role']='admin'
