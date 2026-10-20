// Sheets backup integration with two real signed staff sessions. The private
// harness keeps every spreadsheet row in memory and forbids external HTTP.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
const {spawn,execFileSync}=require('node:child_process'),{once}=require('node:events');
const origin='http://127.0.0.1:8768',sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms));
const harness=`
import copy,importlib,os,threading,requests,uvicorn
os.environ['SUPABASE_URL']='http://supabase.invalid'
os.environ['SUPABASE_SERVICE_ROLE_KEY']=''
def forbidden(*args,**kwargs):raise AssertionError('External provider call forbidden in Sheets backup QA')
requests.sessions.Session.request=forbidden
from tests.gmail_setup_preview import app
from fastapi import HTTPException
from app import auth
from app.utils.gmail_setup import store,sheets,service
importlib.reload(auth)
os.environ['AUTH_DISABLED']='0'
store.team_directory=lambda user:[{'name':name,'email':name.lower()+'@unstuck-agency.com'} for name in ['Dylan','Kelton','Robyn','Chris']]
sheet_rows={};sheet_calls=[];attempts={};held={};gates={}
def await_gate(name):
    held[name]=True
    if not gates.setdefault(name,threading.Event()).wait(20):raise HTTPException(503,'Synthetic backup gate timed out.')
def result(job,live,missing=False):
    gid,title=sheets.TABS[job['section']]
    if missing:return {'live':False,'missing':True,'message':'No sheet entry to update.'}
    row=sheet_rows[job['id']]['row']
    return {'tab':title,'sheet_id':gid,'row':row,'url':sheets.SOURCE_URL+f'#gid={gid}&range=A{row}','live':live}
def publish(job,owner,guard=None):
    key=job['id'];attempts[key]=attempts.get(key,0)+1
    sheet_calls.append({'id':key,'operation':'activate','email':job['email'],'display':job.get('sheet_display_email',job['email']),'owner':owner})
    if job['domain'] in ('slowbackupqa.com','racebackupqa.com') and attempts[key]==1:await_gate(job['domain'])
    if guard:guard()
    if job['domain']=='retrybackupqa.com' and attempts[key]==1:raise HTTPException(503,'Synthetic sheet error: fix row 77, then Retry backup.')
    if key not in sheet_rows:sheet_rows[key]={'row':max([row['row'] for row in sheet_rows.values()]+[19])+1,'email':job.get('sheet_display_email',job['email']),'owner':owner,'order':job['order'],'recovery':'preserve-existing-recovery'}
    sheet_rows[key]['live']=True
    return result(job,True)
def deactivate(job,owner,guard=None):
    key=job['id'];sheet_calls.append({'id':key,'operation':'deactivate','email':job['email'],'owner':owner})
    if job['domain']=='slowbackupqa.com':await_gate('deactivate:'+job['domain'])
    if guard:guard()
    if key not in sheet_rows:return result(job,False,True)
    sheet_rows[key]['live']=False
    return result(job,False)
sheets.publish=publish;sheets.deactivate=deactivate
sheets.readiness=lambda:{'ready':True,'message':'Synthetic sheet write access ready.'}
# Persist a realistic pre-backup client before serving: legacy rows must wait
# for the deliberate client backfill, rather than a historical sweep on load.
seed={'email':'kelton@unstuck-agency.com','name':'Kelton','role':'sdr','tools':['gmail_set']}
def mutate(action,payload):
    record=service.board(seed)
    return service.mutate(seed,record['revision'],action,payload)
member=service.board(seed)['state']['members'][1]['id']
record=mutate('client_save',{'name':'Legacy backup client','section':'client','account_manager_id':member,'operator_id':member})
legacy=record['state']['clients'][-1]['id']
record=mutate('setup_bulk',{'client_id':legacy,'domains':'legacybackupliveqa.com\\nlegacybackupdeletedqa.com\\nlegacypendingbackupqa.com','sender_name':'me','operator_id':member})
legacy_jobs=record['state']['setups'][-3:]
for job in legacy_jobs:mutate('setup_stage',{'id':job['id'],'stage':'Setup complete'})
mutate('setup_stage',{'id':legacy_jobs[1]['id'],'stage':'Gmail deleted','confirmed_auto_renew_stopped':True})
record=service.board(seed)
for job in record['state']['setups']:
    job['sync']={'status':'pending' if job['domain']=='legacypendingbackupqa.com' else 'manual'}
    job.pop('sheet_intent',None);job.pop('sheet_identity',None)
    sheet_rows[job['id']]={'row':10+len(sheet_rows),'email':job['email'],'owner':'Legacy backup client','order':job['order'],'live':job['stage']=='Gmail deleted','recovery':'keep-legacy-recovery'}
store.save(seed,record['state'],record['revision'])
@app.get('/qa/sheets')
def observe_sheets():return {'rows':copy.deepcopy(sheet_rows),'calls':copy.deepcopy(sheet_calls),'attempts':copy.deepcopy(attempts),'held':dict(held)}
@app.post('/qa/sheets/release/{name:path}')
def release_sheet(name:str):gates.setdefault(name,threading.Event()).set();return {'released':name}
uvicorn.run(app,host='127.0.0.1',port=8768,log_level='warning',lifespan='off')
`;
const cookie=name=>execFileSync('.venv/bin/python',['-c',"import base64,json,sys;from itsdangerous import TimestampSigner;print(TimestampSigner('gmail-local-qa').sign(base64.b64encode(json.dumps(json.loads(sys.argv[1])).encode())).decode())",JSON.stringify({user:{email:name.toLowerCase()+'@unstuck-agency.com',name,role:'sdr',tools:['gmail_set']},domain_csrf:name.toLowerCase()+'-sheets-csrf'})],{encoding:'utf8'}).trim();
(async()=>{
 let server,browser,log='';
 try{
  server=spawn('.venv/bin/python',['-c',harness],{stdio:['ignore','pipe','pipe']});server.stdout.on('data',chunk=>log+=chunk);server.stderr.on('data',chunk=>log+=chunk);
  let ready=false;for(let i=0;i<100;i++){if(server.exitCode!==null)throw Error(log);try{ready=(await fetch(origin+'/healthz')).ok;}catch{}if(ready)break;await sleep(100);}assert.equal(ready,true,log);
  browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});const contexts=await Promise.all(['Kelton','Robyn'].map(async name=>{const context=await browser.newContext({viewport:{width:1440,height:1050}});await context.addCookies([{name:'dashboard_session',value:cookie(name),url:origin,httpOnly:true,sameSite:'Lax'}]);return context;}));
  const [kelton,robyn]=await Promise.all(contexts.map(context=>context.newPage())),errors=[],dnsRequests=[];
  for(const page of [kelton,robyn]){page.on('pageerror',error=>errors.push(error.message));page.on('request',request=>{if(/\/api\/gmail-setup\/(?:setups\/[^/]+\/dns-setup|clients\/[^/]+\/(?:dns-setup|dmarc))/.test(request.url()))dnsRequests.push(request.url());});await page.route('**/*',route=>route.request().url().startsWith(origin)?route.continue():route.abort());}
  const idle=page=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
  const data=page=>page.evaluate(()=>fetch('/api/gmail-setup').then(response=>response.json()));
  const observe=()=>kelton.evaluate(()=>fetch('/qa/sheets').then(response=>response.json()));
  const release=name=>kelton.evaluate(name=>fetch('/qa/sheets/release/'+encodeURIComponent(name),{method:'POST'}).then(response=>response.json()),name);
  const change=(page,action,payload)=>page.evaluate(async({action,payload})=>{const board=await fetch('/api/gmail-setup').then(response=>response.json()),response=await fetch('/api/gmail-setup/change',{method:'POST',headers:{'Content-Type':'application/json','X-Domain-CSRF':document.querySelector('#gs-app').dataset.csrf},body:JSON.stringify({revision:board.revision,action,data:payload})});const result=await response.json();if(!response.ok)throw Error(JSON.stringify(result));return result;},{action,payload});
  const open=async(page,id)=>{await page.goto(origin+`/gmail-set?sheetsqa=${Date.now()}${id?'#client='+id:''}`);await idle(page);};
  const card=(page,id)=>page.locator(`[data-setup="${id}"]`),stored=(model,id)=>model.state.setups.find(job=>job.id===id);
  const stage=async(page,id,value)=>{await page.selectOption(`[data-set-stage="${id}"]`,value);await idle(page);await page.click('#gs-client-heading');};
  const waitSync=async(page,id,live)=>{await page.click('#gs-client-heading');await page.waitForFunction(async({id,live})=>{const model=await fetch('/api/gmail-setup').then(response=>response.json()),sync=model.state.setups.find(job=>job.id===id)?.sync;return sync?.status==='synced'&&sync.live===live;},{id,live});await page.click('#gs-refresh');await idle(page);};
  const deleted=async(page,id)=>{await page.selectOption(`[data-set-stage="${id}"]`,'Gmail deleted');await page.waitForFunction(()=>document.querySelector('#gs-renew-dialog').open);const closed=page.evaluate(()=>new Promise(resolve=>document.querySelector('#gs-renew-dialog').addEventListener('close',()=>resolve(true),{once:true})));await page.click('#gs-renew-yes');await closed;await idle(page);await page.click('#gs-client-heading');};
  await open(kelton);let model=await data(kelton);const members=Object.fromEntries(model.state.members.map(member=>[member.name,member.id])),legacy=model.state.clients.find(client=>client.name==='Legacy backup client');
  assert.deepEqual((await observe()).calls,[],'opening the tool must not sweep historical manual records');
  model=await change(kelton,'client_save',{name:'Current Sheets backup client',section:'client',account_manager_id:members.Kelton,operator_id:members.Robyn});const client=model.state.clients.at(-1);
  model=await change(kelton,'setup_bulk',{client_id:client.id,domains:'slowbackupqa.com\nchildbackupqa.com\nretrybackupqa.com\nneverbackupqa.com\nracebackupqa.com',sender_name:'me',operator_id:members.Robyn});const jobs=model.state.setups.filter(job=>job.client_id===client.id),[root,child,retry,never,race]=jobs;
  await change(kelton,'setup_master',{id:root.id,master_domain_id:root.id});await change(kelton,'setup_master',{id:child.id,master_domain_id:root.id});await open(kelton,client.id);await stage(kelton,root.id,'Started');assert.deepEqual((await observe()).calls,[],'unfinished domains never publish');
  // Hold the synthetic write itself. The stage save, sorting and another
  // teammate's edits must finish while Google is still unavailable.
  await stage(kelton,root.id,'Setup complete');await kelton.waitForFunction(async()=>!!(await fetch('/qa/sheets').then(response=>response.json())).held['slowbackupqa.com']);assert.match(await card(kelton,root.id).getAttribute('class'),/gs-domain-complete/);assert.match(await card(kelton,root.id).locator('[data-sheet-status]').textContent(),/Backing up/);assert.equal(await kelton.locator('[data-setup]').last().getAttribute('data-setup'),root.id);assert.equal((await observe()).rows[root.id],undefined);
  await open(robyn,client.id);await stage(robyn,child.id,'Setup complete');await robyn.fill(`[data-set-notes="${child.id}"]`,'Saved by Robyn while a spreadsheet write is paused.');await robyn.click(`[data-save-notes="${child.id}"]`);await idle(robyn);assert.equal(stored(await data(robyn),child.id).notes,'Saved by Robyn while a spreadsheet write is paused.');assert.equal((await observe()).rows[root.id],undefined);// A worker-only revision change is safe to rebase for a client form. Hold
  // its POST until the paused provider completes, forcing an actual stale CAS.
  await open(kelton,client.id);await kelton.click('#gs-edit-client');await kelton.fill('#gs-client-url','https://sheets-client.example.com');let allowClientPost,clientPostReached;const postGate=new Promise(resolve=>allowClientPost=resolve),postReached=new Promise(resolve=>clientPostReached=resolve),clientSaveStatuses=[];
  kelton.on('response',response=>{const request=response.request();if(request.method()==='POST'&&request.url().endsWith('/api/gmail-setup/change')&&request.postDataJSON()?.action==='client_save')clientSaveStatuses.push(response.status());});
  await kelton.route('**/api/gmail-setup/change',async route=>{if(route.request().postDataJSON()?.action==='client_save'){clientPostReached();await postGate;}await route.continue();},{times:1});await kelton.click('#gs-dialog-save');await postReached;await release('slowbackupqa.com');
  await waitSync(robyn,root.id,true);await waitSync(robyn,child.id,true);allowClientPost();await kelton.waitForFunction(()=>!document.querySelector('#gs-dialog').open);await idle(kelton);assert.deepEqual(clientSaveStatuses,[409,200],'a backup-only revision retries the unchanged client form safely');assert.equal((await data(kelton)).state.clients.find(row=>row.id===client.id).website_url,'https://sheets-client.example.com');await waitSync(kelton,root.id,true);let fake=await observe();assert.equal(fake.rows[root.id].email,'me@slowbackupqa.com (Master)');assert.equal(fake.rows[child.id].email,'me@childbackupqa.com (Secondary: slowbackupqa.com is Master)');assert.equal(fake.rows[root.id].live,true);assert.equal(fake.rows[child.id].live,true);assert.notEqual(fake.rows[root.id].row,fake.rows[child.id].row);assert.equal(fake.calls.filter(call=>call.id===root.id&&call.operation==='activate').length,1,'two staff sessions do not duplicate a claimed backup');
  await open(kelton,client.id);assert.equal(await kelton.inputValue(`[data-set-notes="${child.id}"]`),'Saved by Robyn while a spreadsheet write is paused.');assert.match(await card(kelton,root.id).locator('[data-sheet-status]').textContent(),/Backed up · Live/);assert.match(await card(kelton,root.id).locator('[data-sheet-status] a').getAttribute('href'),/#gid=0&range=A\d+/);
  // A real teammate edit still rejects a broad stale form rather than
  // treating it like a harmless backup result or overwriting the new data.
  await kelton.click('#gs-edit-client');await kelton.fill('#gs-client-url','https://must-not-save.example.com');await open(robyn,client.id);await robyn.fill(`[data-set-notes="${child.id}"]`,'A newer teammate note must prevent a broad silent overwrite.');await robyn.click(`[data-save-notes="${child.id}"]`);await idle(robyn);const beforeConflict=clientSaveStatuses.length;await kelton.click('#gs-dialog-save');await idle(kelton);assert.equal(await kelton.locator('#gs-dialog').evaluate(dialog=>dialog.open),true);assert.match(await kelton.textContent('#gs-form-error'),/tracker|updated|changed|refresh/i);assert.deepEqual(clientSaveStatuses.slice(beforeConflict),[409]);assert.equal((await data(kelton)).state.clients.find(row=>row.id===client.id).website_url,'https://sheets-client.example.com');assert.equal(stored(await data(kelton),child.id).notes,'A newer teammate note must prevent a broad silent overwrite.');await kelton.click('#gs-dialog-cancel');await idle(kelton);
  // Cancel leaves Live unchanged. Confirmation immediately changes the local
  // workflow; the later backup only unticks the exact retained sheet row.
  const originalRow=structuredClone(fake.rows[root.id]);await kelton.selectOption(`[data-set-stage="${root.id}"]`,'Gmail deleted');await kelton.waitForFunction(()=>document.querySelector('#gs-renew-dialog').open);await kelton.click('#gs-renew-cancel');await idle(kelton);assert.equal((await observe()).rows[root.id].live,true);await deleted(kelton,root.id);await kelton.waitForFunction(async()=>!!(await fetch('/qa/sheets').then(response=>response.json())).held['deactivate:slowbackupqa.com']);assert.match(await card(kelton,root.id).getAttribute('class'),/gs-domain-deleted/);assert.equal((await observe()).rows[root.id].live,true);await release('deactivate:slowbackupqa.com');await waitSync(kelton,root.id,false);fake=await observe();assert.deepEqual(fake.rows[root.id],{...originalRow,live:false});assert.match(await card(kelton,root.id).locator('[data-sheet-status]').textContent(),/Backed up · Not live/);assert.equal((await data(kelton)).client_progress[client.id].replacements_needed,1);
  await deleted(kelton,never.id);await waitSync(kelton,never.id,false);assert.equal((await observe()).rows[never.id],undefined,'deletion cannot append a nonexistent sheet row');assert.match(await card(kelton,never.id).locator('[data-sheet-status]').textContent(),/No sheet entry to update/);
  // A failed backup is visible to both staff, but never undoes completion.
  await stage(kelton,retry.id,'Setup complete');await kelton.waitForFunction(id=>document.querySelector(`[data-sheet-status="${id}"]`)?.textContent.includes('Backup needs attention'),retry.id);assert.match(await card(kelton,retry.id).getAttribute('class'),/gs-domain-complete/);assert.match(await card(kelton,retry.id).locator('[data-sheet-status]').textContent(),/row 77/);assert.equal(await card(kelton,retry.id).locator('[data-sync]').textContent(),'Retry backup');await open(robyn,client.id);assert.match(await card(robyn,retry.id).locator('[data-sheet-status]').textContent(),/Backup needs attention/);await robyn.click(`[data-sync="${retry.id}"]`);await idle(robyn);await waitSync(robyn,retry.id,true);fake=await observe();assert.equal(fake.attempts[retry.id],2);assert.equal(fake.rows[retry.id].live,true);assert.equal(stored(await data(robyn),retry.id).stage,'Setup complete');
  // A newer Gmail-deleted decision defeats an older in-flight activation,
  // including across devices; the stale writer checks its generation guard.
  await open(kelton,client.id);await stage(kelton,race.id,'Setup complete');await kelton.waitForFunction(async()=>!!(await fetch('/qa/sheets').then(response=>response.json())).held['racebackupqa.com']);await open(robyn,client.id);await deleted(robyn,race.id);await release('racebackupqa.com');await waitSync(robyn,race.id,false);assert.equal((await observe()).rows[race.id],undefined,'a superseded activation must not resurrect Live');assert.equal(stored(await data(robyn),race.id).stage,'Gmail deleted');
  // A domain-only plan has no sheet identity to lock: deleting its unused
  // Gmail creates no row, and reopening must still allow adding the real email.
  model=await change(kelton,'setup_bulk',{client_id:client.id,domains:'blankemailbackupqa.com',sender_name:'',operator_id:members.Robyn});const blank=model.state.setups.at(-1);assert.equal(blank.email,'');await open(kelton,client.id);await deleted(kelton,blank.id);await waitSync(kelton,blank.id,false);model=await data(kelton);assert.equal(stored(model,blank.id).sync.missing,true);assert.equal(stored(model,blank.id).sheet_identity_locked,false);assert.equal((await observe()).rows[blank.id],undefined);await stage(kelton,blank.id,'Not started');assert.equal(await card(kelton,blank.id).locator('[data-set-email]').isEnabled(),true);await card(kelton,blank.id).locator('[data-edit-setup]').click();assert.equal(await kelton.locator('#gs-job-email').isEnabled(),true);await kelton.click('#gs-dialog-cancel');await kelton.fill(`[data-set-email="${blank.id}"]`,'actual@blankemailbackupqa.com');await kelton.click('#gs-client-heading');await idle(kelton);assert.equal(stored(await data(kelton),blank.id).email,'actual@blankemailbackupqa.com');await stage(kelton,blank.id,'Setup complete');await waitSync(kelton,blank.id,true);assert.equal((await observe()).rows[blank.id].email,'actual@blankemailbackupqa.com');
  // Client backfill is deliberate, client-scoped and idempotent for legacy
  // records. Deactivation preserves the historical email/row/recovery data.
  await open(kelton,legacy.id);const legacyJobs=(await data(kelton)).state.setups.filter(job=>job.client_id===legacy.id),beforeLegacy=structuredClone((await observe()).rows);assert.match(await kelton.textContent('#gs-sheet-backup-status'),/not backed up/i);const legacyPending=legacyJobs.find(job=>job.domain==='legacypendingbackupqa.com');assert.doesNotMatch(await card(kelton,legacyPending.id).locator('[data-sheet-status]').textContent(),/Backing up/);assert.equal(await card(kelton,legacyPending.id).locator('[data-sync]').isEnabled(),true,'a legacy pending flag without a durable intent exposes a usable backup action');await kelton.click('#gs-client-sheets');await idle(kelton);for(const job of legacyJobs)await waitSync(kelton,job.id,job.stage==='Setup complete');fake=await observe();for(const job of legacyJobs){assert.equal(fake.rows[job.id].row,beforeLegacy[job.id].row);assert.equal(fake.rows[job.id].email,beforeLegacy[job.id].email);assert.equal(fake.rows[job.id].recovery,'keep-legacy-recovery');assert.equal(fake.rows[job.id].live,job.stage==='Setup complete');}assert.deepEqual(fake.rows[child.id],beforeLegacy[child.id]);const rowCount=Object.keys(fake.rows).length;await kelton.click('#gs-client-sheets');await idle(kelton);for(const job of legacyJobs)await waitSync(kelton,job.id,job.stage==='Setup complete');assert.equal(Object.keys((await observe()).rows).length,rowCount,'repeated backfill cannot duplicate a row');
  // The Biz Dev destination still uses its own authorised strategy tab.
  model=await change(kelton,'client_save',{name:'Biz Dev Sheets backup',section:'bizdev',account_manager_id:members.Kelton,operator_id:members.Robyn});const biz=model.state.clients.at(-1);model=await change(kelton,'setup_bulk',{client_id:biz.id,domains:'bizbackupqa.com',sender_name:'me',operator_id:members.Robyn});const bizJob=model.state.setups.at(-1);await open(robyn,biz.id);await stage(robyn,bizJob.id,'Setup complete');await waitSync(robyn,bizJob.id,true);assert.equal(stored(await data(robyn),bizJob.id).sync.sheet_id,838806676);assert.match(await card(robyn,bizJob.id).locator('[data-sheet-status] a').getAttribute('href'),/#gid=838806676&range=A\d+/);
  await open(robyn,client.id);await card(robyn,retry.id).screenshot({path:'work/gmail-sheets-backed-up-desktop.png',animations:'disabled'});await robyn.setViewportSize({width:390,height:844});assert.equal(await robyn.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false);await card(robyn,root.id).screenshot({path:'work/gmail-sheets-deleted-mobile.png',animations:'disabled'});assert.deepEqual(dnsRequests,[]);assert.deepEqual(errors,[]);console.log('PASS: responsive complete/deleted saves, signed two-staff shared backup status, sequential guarded synthetic writes, Master/Secondary snapshots, Live/Not-live retained rows, missing-row safety, explicit error/retry, safely reopened blank-email plans, legacy pending/client backfill, worker-only CAS retries versus real conflicts, idempotency, Biz Dev routing and mobile.');
 }finally{if(browser)await browser.close();if(server&&server.exitCode===null){for(const name of ['slowbackupqa.com','deactivate:slowbackupqa.com','racebackupqa.com'])try{await fetch(origin+'/qa/sheets/release/'+encodeURIComponent(name),{method:'POST'});}catch{}const exited=once(server,'exit');server.kill('SIGTERM');await Promise.race([exited,sleep(3000)]);if(server.exitCode===null)server.kill('SIGKILL');}}
})().catch(error=>{console.error(error);process.exit(1)});
