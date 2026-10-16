"""Local synthetic tracker harness. Never imported by production."""
import copy
import os
import sys
sys.path.insert(0,os.path.dirname(os.path.dirname(__file__)))
os.environ['AUTH_DISABLED']='1';os.environ['SESSION_SECRET']='gmail-local-qa'
from fastapi import HTTPException
from app.main import app
from app import auth
from app.utils.gmail_setup import store,sheets,dmarc
user={'email':'dylan@unstuck-agency.com','name':'Dylan','role':'admin','tools':['gmail_set']}
auth.get_session_user=lambda request:user
rows={}
store.team_directory=lambda user:[{'name':name,'email':name.lower()+'@unstuck-agency.com'} for name in ['Dylan','Matthew','Robyn','Chris']]
def load(user):return copy.deepcopy(rows.get(store.workspace(user)))
def save(user,state,revision):
    key=store.workspace(user)
    if rows.get(key,{'revision':0})['revision']!=revision:raise HTTPException(409,'Tracker changed. Refresh and retry.')
    from app.utils.gmail_setup.service import now
    rows[key]={'workspace_id':key,'state':copy.deepcopy(state),'revision':revision+1,'updated_at':now()}
    return copy.deepcopy(rows[key])
store.load=load;store.save=save
attempts={}
def publish(job,owner):
    attempts[job['id']]=attempts.get(job['id'],0)+1
    if job['domain']=='retryclients.com' and attempts[job['id']]==1:
        raise HTTPException(503,'Synthetic lost response. Setup saved; Retry sync.')
    gid,title=sheets.TABS[job['section']]
    return {'tab':title,'sheet_id':gid,'row':23,'url':sheets.SOURCE_URL+f'#gid={gid}&range=A23'}
sheets.publish=publish
sheets.readiness=lambda:{'ready':True,'message':'Sheet write access ready.'}

# Synthetic DMARC provider: browser QA cannot call or mutate live registrar DNS.
dmarc_calls=[]
def apply_dmarc(domain,email,key):
    dmarc_calls.append({'domain':domain,'email':email})
    if domain=='conflictdmarc.com':
        raise HTTPException(409,'Existing DMARC differs. Review the record in Porkbun.')
    return {'domain':domain,'email':email,'status':'verified','message':'DMARC verified in synthetic Porkbun readback.','content':f'v=DMARC1; p=none; pct=100; rua=mailto:{email}; ruf=mailto:{email}; fo=1','ttl':600}
dmarc.apply_record=apply_dmarc

# Isolated DNS zones: even browser tests of deletion cannot touch Porkbun.
from app.utils.domain_cluster import provisioning as provisioning
dns_zones={}
dns_deletes=[]
dns_forwards={}
class SyntheticDNS:
    def check(self):return {'credentials_valid':True,'read':True,'dns':True}
    def call(self,method,path,data=None,idempotency=None):
        if path.startswith('/domain/get/'):
            return {'domain':{'domain':path.split('/')[-1],'status':'ACTIVE','apiAccess':1}}
        if path.startswith('/domain/getNs/'):
            return {'ns':['curitiba.ns.porkbun.com']}
        if path.startswith('/dns/retrieve/'):
            domain=path.split('/')[-1]
            if domain not in dns_zones:
                dns_zones[domain]=[{'id':str(i+1),'name':host,'type':kind,'content':content,'ttl':'600','prio':'0'} for i,(host,kind,content) in enumerate([
                    (domain,'MX','fwd1.porkbun.com'),(domain,'MX','fwd2.porkbun.com'),
                    (domain,'TXT','v=spf1 include:_spf.porkbun.com ~all'),
                    ('_acme-challenge.'+domain,'TXT','challenge-one'),('_acme-challenge.'+domain,'TXT','challenge-two'),
                    (domain,'MX','smtp.google.com'),('_dmarc.'+domain,'TXT',provisioning.dmarc_value(domain,{'rua':'me@'+domain,'ruf':'me@'+domain}))])]
            return {'records':copy.deepcopy(dns_zones[domain])}
        if path.startswith('/domain/getUrlForwarding/'):
            return {'forwards':copy.deepcopy(dns_forwards.get(path.split('/')[-1],[]))}
        if path.startswith('/domain/addUrlForward/'):
            domain=path.split('/')[-1];dns_forwards.setdefault(domain,[]).append(copy.deepcopy(data))
            return {'status':'SUCCESS'}
        if path.startswith('/dns/create/'):
            domain=path.split('/')[-1];identifier=str(100+len(dns_zones[domain]))
            dns_zones[domain].append({**data,'id':identifier,'name':data['name']+'.'+domain,'ttl':str(data['ttl']),'prio':'0'})
            return {'status':'SUCCESS','id':identifier}
        if path.startswith('/dns/delete/'):
            assert method=='post' and idempotency
            domain,identifier=path.split('/')[-2:]
            dns_deletes.append((domain,identifier))
            dns_zones[domain]=[r for r in dns_zones[domain] if r['id']!=identifier]
            return {'status':'SUCCESS'}
        raise AssertionError('Unexpected synthetic DNS endpoint')
provisioning.Porkbun=SyntheticDNS
