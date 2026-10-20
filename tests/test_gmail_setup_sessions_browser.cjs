// Runs its own isolated preview with real signed Kelton/Robyn SDR sessions.
// No production credentials, provider requests or production board writes.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
const {spawn,execFileSync}=require('node:child_process'),{once}=require('node:events');
const origin='http://127.0.0.1:8767',sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms));
const harness=`
import importlib,os,requests,uvicorn
os.environ['SUPABASE_URL']='http://supabase.invalid'
os.environ['SUPABASE_SERVICE_ROLE_KEY']=''
def forbidden(*args,**kwargs):raise AssertionError('External provider call forbidden in session QA')
requests.sessions.Session.request=forbidden
from tests.gmail_setup_preview import app
from app import auth
from app.utils.gmail_setup import store
importlib.reload(auth)
os.environ['AUTH_DISABLED']='0'
store.team_directory=lambda user:[{'name':name,'email':name.lower()+'@unstuck-agency.com'} for name in ['Dylan','Kelton','Robyn','Chris']]
uvicorn.run(app,host='127.0.0.1',port=8767,log_level='warning',lifespan='off')
`;
const signedCookie=(name,token,tools=['gmail_set'])=>execFileSync('.venv/bin/python',['-c',
 "import base64,json,sys;from itsdangerous import TimestampSigner;print(TimestampSigner('gmail-local-qa').sign(base64.b64encode(json.dumps(json.loads(sys.argv[1])).encode())).decode())",
 JSON.stringify({user:{email:name.toLowerCase()+'@unstuck-agency.com',name,role:'sdr',tools},domain_csrf:token})],{encoding:'utf8'}).trim();
(async()=>{
 let browser,server,serverLog='';
 try{
  server=spawn('.venv/bin/python',['-c',harness],{stdio:['ignore','pipe','pipe']});
  server.stdout.on('data',data=>serverLog+=data);server.stderr.on('data',data=>serverLog+=data);
  let ready=false;for(let i=0;i<100;i++){if(server.exitCode!==null)throw Error('Synthetic server failed: '+serverLog);try{ready=(await fetch(origin+'/healthz')).ok;}catch{}if(ready)break;await sleep(100);}assert.equal(ready,true,'isolated signed-session server starts');
  browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});
  const contexts=await Promise.all([browser.newContext({viewport:{width:1440,height:1050}}),browser.newContext({viewport:{width:1440,height:1050}})]),[kelton,robyn]=await Promise.all(contexts.map(context=>context.newPage()));
  const errors=[],syncRequests=[];
  for(const page of [kelton,robyn]){page.on('pageerror',e=>errors.push(e.message));page.on('dialog',dialog=>dialog.accept());page.on('request',r=>{if(r.method()==='POST'&&/\/setups\/[^/]+\/sync$/.test(r.url()))syncRequests.push(r.url());});await page.route('**/*',route=>route.request().url().startsWith(origin)?route.continue():route.abort());}
  const setSession=(index,name,token,tools)=>contexts[index].addCookies([{name:'dashboard_session',value:signedCookie(name,token,tools),url:origin,httpOnly:true,sameSite:'Lax'}]);
  await setSession(0,'Kelton','kelton-v1');await setSession(1,'Robyn','robyn-v1');
  const idle=page=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
  const data=page=>page.evaluate(()=>fetch('/api/gmail-setup').then(r=>r.json()));
  const change=(page,action,payload)=>page.evaluate(async({action,payload})=>{const board=await fetch('/api/gmail-setup').then(r=>r.json());const r=await fetch('/api/gmail-setup/change',{method:'POST',headers:{'Content-Type':'application/json','X-Domain-CSRF':document.querySelector('#gs-app').dataset.csrf},body:JSON.stringify({revision:board.revision,action,data:payload})});const result=await r.json();if(!r.ok)throw Error(JSON.stringify(result));return result;},{action,payload});
  const open=async(page,id)=>{await page.goto(origin+`/gmail-set?sessionqa=${Date.now()}${id?'#client='+id:''}`);await idle(page);};
  const writes=page=>{const statuses=[],sessions=[];const listener=response=>{if(response.request().method()==='POST'&&response.url().endsWith('/api/gmail-setup/change'))statuses.push(response.status());if(response.request().method()==='GET'&&response.url().endsWith('/api/gmail-setup/session'))sessions.push(response.status());};page.on('response',listener);return {statuses,sessions,stop:()=>page.off('response',listener)};};
  const stage=async(page,id,value)=>{await page.selectOption(`[data-set-stage="${id}"]`,value);await idle(page);};
  const newerVersion=async(page,marker)=>page.route('**/api/gmail-setup/session',async route=>{const response=await route.fetch(),body=await response.json();await route.fulfill({response,json:{...body,app_version:marker}});},{times:1});
  const detectVersion=async page=>{const response=page.waitForResponse(r=>r.request().method()==='GET'&&r.url().endsWith('/api/gmail-setup/session'));await page.evaluate(()=>document.dispatchEvent(new Event('visibilitychange')));await (await response).finished();};
  const reloadEvent=page=>page.waitForEvent('framenavigated',{predicate:frame=>frame===page.mainFrame()});
  const loaded=async page=>{await page.waitForLoadState('domcontentloaded');await idle(page);};
  await open(kelton);const members=Object.fromEntries((await data(kelton)).state.members.map(m=>[m.name,m.id]));
  let model=await change(kelton,'client_save',{name:'Signed session browser',section:'client',account_manager_id:members.Kelton,operator_id:members.Robyn,domain_target:10,sender_full_name:'Richard Sessions',website_url:'sessions.example.com',account_manager_email:'kelton@unstuck-agency.com'}),client=model.state.clients.at(-1);
  model=await change(kelton,'setup_bulk',{client_id:client.id,domains:'signedsessionone.com\nsignedsessiontwo.com\nsignedsessiondelete.com',sender_name:'richard',operator_id:members.Robyn});const [first,second,removed]=model.state.setups;
  await Promise.all([open(kelton,client.id),open(robyn,client.id)]);
  for(const page of [kelton,robyn]){assert.equal(await page.getAttribute('#gs-app','data-dns-enabled'),'true');assert.equal(await page.locator('#gs-client-dns').isVisible(),true);assert.match(await page.textContent('#gs-client-info'),/Richard Sessions/);assert.match(await page.textContent('#gs-client-info'),/https:\/\/sessions.example.com/);assert.match(await page.textContent('#gs-client-info'),/kelton@unstuck-agency.com/);}
  // A fresh login/session cookie can arrive while an existing tab still holds
  // its previous CSRF token. Recover once without discarding the chosen stage.
  await setSession(0,'Kelton','kelton-v2');const keltonRecovery=writes(kelton);await stage(kelton,first.id,'Started');keltonRecovery.stop();
  assert.deepEqual(keltonRecovery.statuses,[403,200]);assert.equal(keltonRecovery.sessions.length,1);assert.equal(await kelton.getAttribute('#gs-app','data-csrf'),'kelton-v2');assert.equal((await data(kelton)).state.setups.find(j=>j.id===first.id).stage,'Started');
  assert.equal((await data(kelton)).state.events.filter(e=>e.by==='kelton@unstuck-agency.com'&&e.action==='setup_stage').length,1,'a rejected first request never applies the mutation');
  // Robyn's tab also has a stale workspace revision. Token recovery and the
  // normal safe revision retry must work together without duplicating notes.
  await setSession(1,'Robyn','robyn-v2');const robynRecovery=writes(robyn);await robyn.fill(`[data-set-notes="${second.id}"]`,'Saved after renewing the Robyn session.');await robyn.click(`[data-save-notes="${second.id}"]`);await idle(robyn);robynRecovery.stop();
  assert.deepEqual(robynRecovery.statuses,[403,409,200]);assert.equal(robynRecovery.sessions.length,1);model=await data(robyn);assert.equal(model.state.setups.find(j=>j.id===second.id).notes,'Saved after renewing the Robyn session.');assert.equal(model.state.setups.find(j=>j.id===first.id).stage,'Started');
  assert.equal(model.state.events.filter(e=>e.by==='robyn@unstuck-agency.com'&&e.action==='setup_notes').length,1);assert.deepEqual(syncRequests,[]);
  // A running older tab must discover a release without silently discarding
  // a notes draft. Saving it permits the pending upgrade to reload safely.
  let navigations=0;kelton.on('framenavigated',frame=>{if(frame===kelton.mainFrame())navigations++;});
  await kelton.fill(`[data-set-notes="${first.id}"]`,'Keep this draft when a new release appears.');await kelton.evaluate(()=>document.activeElement?.blur());
  await newerVersion(kelton,'qa-release-notes');const beforeDraft=navigations;await detectVersion(kelton);await kelton.locator('#gs-update').waitFor({state:'visible'});
  assert.equal(await kelton.locator('#gs-apply-update').isDisabled(),true);assert.equal(await kelton.inputValue(`[data-set-notes="${first.id}"]`),'Keep this draft when a new release appears.');assert.equal(navigations,beforeDraft,'a release does not reload over unsaved notes');
  let reloaded=reloadEvent(kelton);await kelton.click(`[data-save-notes="${first.id}"]`);await reloaded;await loaded(kelton);assert.equal(await kelton.inputValue(`[data-set-notes="${first.id}"]`),'Keep this draft when a new release appears.');assert.equal(await kelton.locator('#gs-update').isVisible(),false);
  // Likewise, an active write must finish before loading the new release.
  let releaseWrite,capturedWrite;const writeGate=new Promise(resolve=>releaseWrite=resolve),captured=new Promise(resolve=>capturedWrite=resolve);
  await kelton.route('**/api/gmail-setup/change',async route=>{const response=await route.fetch();capturedWrite();await writeGate;await route.fulfill({response});},{times:1});
  await kelton.selectOption(`[data-set-stage="${first.id}"]`,'Purchase made');await captured;
  await newerVersion(kelton,'qa-release-during-save');const beforeWrite=navigations;await detectVersion(kelton);await kelton.locator('#gs-update').waitFor({state:'visible'});assert.equal(await kelton.locator('#gs-apply-update').isDisabled(),true);assert.equal(navigations,beforeWrite,'a pending write prevents a release reload');
  reloaded=reloadEvent(kelton);releaseWrite();await reloaded;await loaded(kelton);assert.equal(await kelton.inputValue(`[data-set-stage="${first.id}"]`),'Purchase made');assert.equal((await data(kelton)).state.setups.find(j=>j.id===first.id).stage,'Purchase made');
  // Refresh checks the release as well as data; idle stale tabs load it fully.
  await newerVersion(kelton,'qa-release-refresh-button');reloaded=reloadEvent(kelton);await kelton.click('#gs-refresh');await reloaded;await loaded(kelton);assert.equal(await kelton.inputValue(`[data-set-stage="${first.id}"]`),'Purchase made');
  // An open client form stays intact until its user explicitly closes it.
  await kelton.click('#gs-edit-client');await kelton.fill('#gs-client-sender-full-name','Unsaved client form draft');await newerVersion(kelton,'qa-release-dialog');const beforeDialog=navigations;await detectVersion(kelton);await kelton.locator('#gs-update').waitFor({state:'visible'});
  assert.equal(await kelton.locator('#gs-dialog').evaluate(el=>el.open),true);assert.equal(await kelton.inputValue('#gs-client-sender-full-name'),'Unsaved client form draft');assert.equal(await kelton.locator('#gs-apply-update').isDisabled(),true);assert.equal(navigations,beforeDialog);
  reloaded=reloadEvent(kelton);await kelton.click('#gs-dialog-cancel');await reloaded;await loaded(kelton);assert.match(await kelton.textContent('#gs-client-info'),/Richard Sessions/);
  // Deleting the domain also discards that domain's unsaved note draft. It
  // must not leave a hidden orphan draft blocking every later upgrade.
  await kelton.fill(`[data-set-notes="${removed.id}"]`,'Draft belonging to a domain we will delete.');await kelton.evaluate(()=>document.activeElement?.blur());await newerVersion(kelton,'qa-release-delete-draft');const beforeDelete=navigations;await detectVersion(kelton);await kelton.locator('#gs-update').waitFor({state:'visible'});assert.equal(await kelton.locator('#gs-apply-update').isDisabled(),true);assert.equal(navigations,beforeDelete);
  reloaded=reloadEvent(kelton);await kelton.click(`[data-remove-setup="${removed.id}"]`);await reloaded;await loaded(kelton);assert.equal(await kelton.locator(`[data-setup="${removed.id}"]`).count(),0);assert.equal((await data(kelton)).state.setups.find(j=>j.id===removed.id).archived,true);assert.equal(await kelton.locator('#gs-update').isVisible(),false);
  // Only a definitive pre-write session failure can be replayed, and only
  // once. A persistent failure stops; an authorization denial never retries.
  const beforeFailure=(await data(kelton)).revision;await kelton.route('**/api/gmail-setup/change',route=>route.fulfill({status:403,headers:{'Content-Type':'application/json','X-Gmail-Session':'refresh'},json:{detail:'Synthetic persistent session verification failure.'}}));
  const bounded=writes(kelton);await stage(kelton,first.id,'Not started');bounded.stop();assert.deepEqual(bounded.statuses,[403,403]);assert.equal(bounded.sessions.length,1);assert.equal((await data(kelton)).revision,beforeFailure);assert.equal(await kelton.inputValue(`[data-set-stage="${first.id}"]`),'Purchase made');assert.match(await kelton.textContent('#gs-message'),/session verification failure/i);
  await setSession(1,'Robyn','robyn-denied',[]);const denied=writes(robyn);await robyn.fill(`[data-set-notes="${second.id}"]`,'This authorization-denied edit must not save.');await robyn.click(`[data-save-notes="${second.id}"]`);await idle(robyn);denied.stop();assert.deepEqual(denied.statuses,[403]);assert.equal(denied.sessions.length,0);assert.equal((await data(kelton)).state.setups.find(j=>j.id===second.id).notes,'Saved after renewing the Robyn session.');
  assert.deepEqual(syncRequests,[]);assert.deepEqual(errors,[]);console.log('PASS: genuine Kelton/Robyn SDR sessions, one-time CSRF recovery plus safe CAS retry, shared stages/client metadata, new-release detection, draft/dialog/write protection, automatic safe upgrades, Refresh release checks, bounded failures, authorization denials and zero sheet writes.');
 }finally{if(browser)await browser.close();if(server&&server.exitCode===null){const exited=once(server,'exit');server.kill('SIGTERM');await Promise.race([exited,sleep(3000)]);}}
})().catch(e=>{console.error(e);process.exit(1)});
