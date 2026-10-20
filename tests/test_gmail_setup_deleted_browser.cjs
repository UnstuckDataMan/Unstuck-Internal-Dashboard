// Gmail-deleted workflow regression with real signed local sessions. All
// persistence/providers are synthetic; confirming never changes live Porkbun.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
const {spawn,execFileSync}=require('node:child_process'),{once}=require('node:events');
const origin='http://127.0.0.1:8768',sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms));
const harness=`
import importlib,os,requests,uvicorn
os.environ['SUPABASE_URL']='http://supabase.invalid'
os.environ['SUPABASE_SERVICE_ROLE_KEY']=''
def forbidden(*args,**kwargs):raise AssertionError('External provider call forbidden in deleted-workflow QA')
requests.sessions.Session.request=forbidden
from tests.gmail_setup_preview import app
from app import auth
from app.utils.gmail_setup import store
importlib.reload(auth)
os.environ['AUTH_DISABLED']='0'
store.team_directory=lambda user:[{'name':name,'email':name.lower()+'@unstuck-agency.com'} for name in ['Dylan','Kelton','Robyn','Chris']]
uvicorn.run(app,host='127.0.0.1',port=8768,log_level='warning',lifespan='off')
`;
const cookie=(name,role='sdr')=>execFileSync('.venv/bin/python',['-c',"import base64,json,sys;from itsdangerous import TimestampSigner;print(TimestampSigner('gmail-local-qa').sign(base64.b64encode(json.dumps(json.loads(sys.argv[1])).encode())).decode())",JSON.stringify({user:{email:name.toLowerCase()+'@unstuck-agency.com',name,role,tools:['gmail_set']},domain_csrf:name.toLowerCase()+'-deleted-csrf'})],{encoding:'utf8'}).trim();
(async()=>{
 let server,browser,log='';
 try{
  server=spawn('.venv/bin/python',['-c',harness],{stdio:['ignore','pipe','pipe']});server.stdout.on('data',data=>log+=data);server.stderr.on('data',data=>log+=data);
  let ready=false;for(let i=0;i<100;i++){if(server.exitCode!==null)throw Error(log);try{ready=(await fetch(origin+'/healthz')).ok;}catch{}if(ready)break;await sleep(100);}assert.equal(ready,true);
  browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});const contexts=await Promise.all([browser.newContext({viewport:{width:1440,height:1050}}),browser.newContext({viewport:{width:1440,height:1050}}),browser.newContext({viewport:{width:1440,height:1050}})]);
  for(const [index,name] of ['Kelton','Robyn','Dylan'].entries())await contexts[index].addCookies([{name:'dashboard_session',value:cookie(name,index===2?'admin':'sdr'),url:origin,httpOnly:true,sameSite:'Lax'}]);
  const [kelton,robyn,admin]=await Promise.all(contexts.map(context=>context.newPage())),errors=[],providerRequests=[],posted=[];
  for(const page of [kelton,robyn,admin]){page.on('pageerror',error=>errors.push(error.message));page.on('request',request=>{if(request.method()==='POST'&&/\/api\/gmail-setup\/(?:setups\/[^/]+\/(?:sync|dns-setup)|clients\/[^/]+\/(?:dns-setup|dmarc))/.test(request.url()))providerRequests.push(request.url());if(request.method()==='POST'&&request.url().endsWith('/api/gmail-setup/change'))posted.push(request.postDataJSON());});await page.route('**/*',route=>route.request().url().startsWith(origin)?route.continue():route.abort());}
  const idle=page=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
  const data=page=>page.evaluate(()=>fetch('/api/gmail-setup').then(response=>response.json()));
  const requestChange=(page,action,payload)=>page.evaluate(async({action,payload})=>{const board=await fetch('/api/gmail-setup').then(response=>response.json()),response=await fetch('/api/gmail-setup/change',{method:'POST',headers:{'Content-Type':'application/json','X-Domain-CSRF':document.querySelector('#gs-app').dataset.csrf},body:JSON.stringify({revision:board.revision,action,data:payload})});return {status:response.status,result:await response.json()};},{action,payload});
  const change=async(page,action,payload)=>{const response=await requestChange(page,action,payload);assert.equal(response.status,200,JSON.stringify(response.result));return response.result;};
  const open=async(page,id)=>{await page.goto(origin+`/gmail-set?deletedqa=${Date.now()}${id?'#client='+id:''}`);await idle(page);};
  const card=(page,id)=>page.locator(`[data-setup="${id}"]`),stored=(model,id)=>model.state.setups.find(job=>job.id===id);
  const stage=async(id,value)=>{await kelton.selectOption(`[data-set-stage="${id}"]`,value);await idle(kelton);};
  const prompt=async(id)=>{await kelton.selectOption(`[data-set-stage="${id}"]`,'Gmail deleted');await kelton.waitForFunction(()=>document.querySelector('#gs-renew-dialog').open);};
  const closeRenew=async action=>{const closed=kelton.evaluate(()=>new Promise(resolve=>document.querySelector('#gs-renew-dialog').addEventListener('close',()=>resolve(true),{once:true})));await action();await closed;await idle(kelton);};
  const yes=()=>closeRenew(()=>kelton.click('#gs-renew-yes'));
  const deleteStage=async id=>{await prompt(id);await yes();await kelton.waitForFunction(id=>document.querySelector(`[data-set-stage="${id}"]`)?.value==='Gmail deleted',id);};
  await open(kelton);const members=Object.fromEntries((await data(kelton)).state.members.map(member=>[member.name,member.id]));
  let model=await change(kelton,'devices_add',{member_id:members.Robyn,phones:2}),phones=model.state.devices;
  model=await change(kelton,'client_save',{name:'Gmail deleted workflow',section:'client',account_manager_id:members.Kelton,operator_id:members.Robyn,domain_target:7});const client=model.state.clients.at(-1);
  model=await change(kelton,'setup_bulk',{client_id:client.id,domains:Array.from({length:7},(_,index)=>`deletedqa${index}.com`).join('\n'),sender_name:'me',operator_id:members.Robyn});const jobs=model.state.setups;
  await change(kelton,'setup_assign',{id:jobs[0].id,device_id:phones[0].id});await change(kelton,'setup_notes',{id:jobs[0].id,notes:'Retain completed purchase history.'});await change(kelton,'setup_master',{id:jobs[0].id,master_domain_id:jobs[0].id});await change(kelton,'setup_master',{id:jobs[2].id,master_domain_id:jobs[0].id});
  await change(kelton,'setup_stage',{id:jobs[0].id,stage:'Setup complete'});await change(kelton,'setup_assign',{id:jobs[1].id,device_id:phones[1].id});await change(kelton,'setup_stage',{id:jobs[1].id,stage:'Purchase made'});await change(kelton,'setup_stage',{id:jobs[1].id,stage:'Burned'});await change(kelton,'setup_stage',{id:jobs[4].id,stage:'Setup complete'});model=await change(kelton,'setup_stage',{id:jobs[5].id,stage:'Burned'});
  const historyComplete=structuredClone(stored(model,jobs[0].id)),historyBurned=structuredClone(stored(model,jobs[1].id)),uses=model.state.devices.map(device=>device.use_count);
  await open(kelton,client.id);
  // Cancelling or pressing Escape never changes the domain, and the row's
  // selected stage is restored while the confirmation question is open.
  let before=await data(kelton);await prompt(jobs[2].id);assert.match(await kelton.textContent('#gs-renew-dialog'),/Have you stopped domain auto-renew\?/);assert.equal(await kelton.inputValue(`[data-set-stage="${jobs[2].id}"]`),'Not started');assert.equal((await data(kelton)).revision,before.revision);await kelton.locator('#gs-renew-dialog').screenshot({path:'work/gmail-deleted-confirmation.png',animations:'disabled'});
  await closeRenew(()=>kelton.click('#gs-renew-cancel'));assert.equal((await data(kelton)).revision,before.revision);assert.equal(await kelton.inputValue(`[data-set-stage="${jobs[2].id}"]`),'Not started');
  await prompt(jobs[2].id);await closeRenew(()=>kelton.keyboard.press('Escape'));assert.equal((await data(kelton)).revision,before.revision);
  for(const flag of [undefined,false]){const rejected=await requestChange(kelton,'setup_stage',{id:jobs[2].id,stage:'Gmail deleted',...(flag===undefined?{}:{confirmed_auto_renew_stopped:flag})});assert.ok([400,409].includes(rejected.status));assert.match(JSON.stringify(rejected.result),/auto.?renew|confirm/i);assert.equal((await data(kelton)).revision,before.revision);}
  await deleteStage(jobs[2].id);assert.equal(stored(await data(kelton),jobs[2].id).archived,false);
  const inlineRequest=posted.findLast(request=>request.action==='setup_stage'&&request.data.id===jobs[2].id);assert.equal(inlineRequest.data.confirmed_auto_renew_stopped,true);
  // Complete and burned domains can become deleted without losing historical
  // completion order, purchase/device use, notes or master relationships.
  await deleteStage(jobs[0].id);await deleteStage(jobs[1].id);model=await data(kelton);
  for(const key of ['order','completed_at','purchase_at','purchase_device','notes','master_domain_id'])assert.deepEqual(stored(model,jobs[0].id)[key],historyComplete[key]);for(const key of ['burned_at','burned_from','purchase_at','purchase_device'])assert.deepEqual(stored(model,jobs[1].id)[key],historyBurned[key]);assert.deepEqual(model.state.devices.map(device=>device.use_count),uses);
  // Details also asks immediately. Yes only selects the stage in that form;
  // the record is unchanged until Save. Cancelling the form resets consent.
  await card(kelton,jobs[3].id).locator('[data-edit-setup]').click();await kelton.fill('#gs-job-notes','Details draft remains while confirming.');before=await data(kelton);await kelton.selectOption('#gs-job-stage','Gmail deleted');await kelton.waitForFunction(()=>document.querySelector('#gs-renew-dialog').open);await closeRenew(()=>kelton.click('#gs-renew-cancel'));assert.equal(await kelton.inputValue('#gs-job-stage'),'Not started');assert.equal(await kelton.inputValue('#gs-job-notes'),'Details draft remains while confirming.');
  await kelton.selectOption('#gs-job-stage','Gmail deleted');await kelton.waitForFunction(()=>document.querySelector('#gs-renew-dialog').open);await yes();await kelton.waitForFunction(()=>document.querySelector('#gs-job-stage').value==='Gmail deleted');assert.equal((await data(kelton)).revision,before.revision);await kelton.click('#gs-dialog-cancel');assert.equal((await data(kelton)).revision,before.revision);
  await card(kelton,jobs[3].id).locator('[data-edit-setup]').click();await kelton.fill('#gs-job-notes','Saved when the Details form completes deletion.');await kelton.selectOption('#gs-job-stage','Gmail deleted');await kelton.waitForFunction(()=>document.querySelector('#gs-renew-dialog').open);await yes();await kelton.click('#gs-dialog-save');await kelton.waitForFunction(()=>!document.querySelector('#gs-dialog').open);await idle(kelton);assert.equal(stored(await data(kelton),jobs[3].id).stage,'Gmail deleted');assert.equal(posted.findLast(request=>request.action==='setup_save'&&request.data.id===jobs[3].id).data.confirmed_auto_renew_stopped,true);
  // Deleted records remain visible and editable for notes/grouping, with the
  // account-setup fields locked and their former setup capacity still used.
  for(const job of jobs.slice(0,4)){const node=card(kelton,job.id);assert.match(await node.getAttribute('class'),/gs-domain-deleted/);for(const field of ['email','operator','device'])assert.equal(await node.locator(`[data-set-${field}]`).isDisabled(),true);assert.equal(await node.locator('[data-set-notes]').isEnabled(),true);assert.equal(await node.locator('[data-set-master]').isEnabled(),true);}
  await kelton.fill(`[data-set-notes="${jobs[0].id}"]`,'Deleted Gmail: note and master history remain saved.');await kelton.click(`[data-save-notes="${jobs[0].id}"]`);await idle(kelton);await kelton.selectOption(`[data-set-master="${jobs[1].id}"]`,jobs[1].id);await idle(kelton);await kelton.selectOption(`[data-set-master="${jobs[2].id}"]`,jobs[1].id);await idle(kelton);
  await open(robyn,client.id);assert.equal(await robyn.inputValue(`[data-set-notes="${jobs[0].id}"]`),'Deleted Gmail: note and master history remain saved.');assert.equal(await card(robyn,jobs[2].id).getAttribute('data-master-root'),jobs[1].id);await robyn.reload();await idle(robyn);assert.equal(await robyn.inputValue(`[data-set-stage="${jobs[0].id}"]`),'Gmail deleted');assert.equal(await card(robyn,jobs[0].id).locator('svg[aria-label="Master domain"]').count(),1);
  const headings=await kelton.locator('#gs-setup-list > h4').allTextContents();assert.ok(headings.findIndex(text=>text.startsWith('Live domains (complete)'))<headings.findIndex(text=>text.startsWith('Burned')));assert.ok(headings.findIndex(text=>text.startsWith('Burned'))<headings.findIndex(text=>text.startsWith('Deleted domains')));
  await kelton.selectOption('#gs-filter-stage','Gmail deleted');assert.equal(await kelton.locator('[data-setup]').count(),4);assert.equal(await kelton.locator('#gs-setup-list > h4').count(),1);assert.match(await kelton.textContent('#gs-setup-list > h4'),/^Deleted domains/);await kelton.selectOption('#gs-filter-stage','');
  model=await data(kelton);assert.equal(model.client_progress[client.id].target,7);assert.equal(model.client_progress[client.id].complete,1);assert.equal(model.client_progress[client.id].remaining,6);assert.equal(model.summary.todo,6);
  // Bulk completion skips every deleted and burned row. The DNS bulk dialog
  // likewise previews only the two remaining live domains, without applying.
  await kelton.click('#gs-complete-client');assert.match(await kelton.textContent('#gs-dialog'),/1 unfinished domain/);await kelton.click('#gs-dialog-save');await kelton.waitForFunction(()=>!document.querySelector('#gs-dialog').open);await idle(kelton);model=await data(kelton);for(const job of jobs.slice(0,4))assert.equal(stored(model,job.id).stage,'Gmail deleted');assert.equal(stored(model,jobs[5].id).stage,'Burned');assert.equal(stored(model,jobs[6].id).stage,'Setup complete');assert.equal(model.client_progress[client.id].remaining,5);assert.equal(model.client_progress[client.id].replacements_needed,5);
  await open(admin,client.id);for(const job of jobs.slice(0,4))assert.equal(await card(admin,job.id).locator('[data-setup-dns]').count(),0);assert.equal(await card(admin,jobs[4].id).locator('[data-setup-dns]').count(),1);await admin.click('#gs-client-dns');const preview=await admin.locator('.gs-dmarc-preview li').allTextContents();assert.equal(preview.length,2);for(const job of jobs.slice(0,4))assert.ok(preview.every(text=>!text.includes(job.domain)));assert.ok(preview.some(text=>text.includes(jobs[4].domain)));assert.ok(preview.some(text=>text.includes(jobs[6].domain)));await admin.click('#gs-dialog-cancel');
  await card(kelton,jobs[0].id).screenshot({path:'work/gmail-deleted-card-desktop.png',animations:'disabled'});await kelton.setViewportSize({width:390,height:844});assert.equal(await kelton.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false);await card(kelton,jobs[0].id).screenshot({path:'work/gmail-deleted-card-mobile.png',animations:'disabled'});
  assert.deepEqual(providerRequests,[]);assert.deepEqual(errors,[]);console.log('PASS: Gmail-deleted confirmation Cancel/Escape/Yes, guarded API and Details saves, preserved complete/burned/device/master history, staff notes/reload, separate ordered sections/filter, bulk exclusions, replacement demand, admin DNS preview exclusion, mobile and zero provider mutations.');
 }finally{if(browser)await browser.close();if(server&&server.exitCode===null){const exited=once(server,'exit');server.kill('SIGTERM');await Promise.race([exited,sleep(3000)]);}}
})().catch(error=>{console.error(error);process.exit(1)});
