// Synthetic password-sharing regression. Only local QA passwords are used;
// genuine signed staff sessions exercise the normal auth/workspace controls.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
const {spawn,execFileSync}=require('node:child_process'),{once}=require('node:events');
const origin='http://127.0.0.1:8768',sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms));
const harness=`
import importlib,os,requests,uvicorn
os.environ['SUPABASE_URL']='http://supabase.invalid'
os.environ['SUPABASE_SERVICE_ROLE_KEY']=''
def forbidden(*args,**kwargs):raise AssertionError('External provider call forbidden in password QA')
requests.sessions.Session.request=forbidden
from tests.gmail_setup_preview import app
from app import auth
from app.utils.gmail_setup import store
importlib.reload(auth)
os.environ['AUTH_DISABLED']='0'
store.team_directory=lambda user:[{'name':name,'email':name.lower()+'@unstuck-agency.com'} for name in ['Dylan','Kelton','Robyn','Chris']]
uvicorn.run(app,host='127.0.0.1',port=8768,log_level='warning',lifespan='off')
`;
const cookie=(name,{email=name.toLowerCase()+'@unstuck-agency.com',tools=['gmail_set'],token=name.toLowerCase()+'-password-csrf'}={})=>execFileSync('.venv/bin/python',['-c',"import base64,json,sys;from itsdangerous import TimestampSigner;print(TimestampSigner('gmail-local-qa').sign(base64.b64encode(json.dumps(json.loads(sys.argv[1])).encode())).decode())",JSON.stringify({user:{email,name,role:'sdr',tools},domain_csrf:token})],{encoding:'utf8'}).trim();
(async()=>{
 let server,browser,log='';
 try{
  server=spawn('.venv/bin/python',['-c',harness],{stdio:['ignore','pipe','pipe']});server.stdout.on('data',data=>log+=data);server.stderr.on('data',data=>log+=data);
  let ready=false;for(let i=0;i<100;i++){if(server.exitCode!==null)throw Error(log);try{ready=(await fetch(origin+'/healthz')).ok;}catch{}if(ready)break;await sleep(100);}assert.equal(ready,true);
  browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});const contexts=await Promise.all([browser.newContext({viewport:{width:1440,height:1050}}),browser.newContext({viewport:{width:1440,height:1050}})]);
  for(const [index,name] of ['Kelton','Robyn'].entries())await contexts[index].addCookies([{name:'dashboard_session',value:cookie(name),url:origin,httpOnly:true,sameSite:'Lax'}]);
  const [kelton,robyn]=await Promise.all(contexts.map(context=>context.newPage())),errors=[],providerRequests=[],reveals=[];
  for(const page of [kelton,robyn]){page.on('pageerror',error=>errors.push(error.message));page.on('dialog',dialog=>dialog.accept());page.on('request',request=>{if(request.method()==='POST'&&/\/api\/gmail-setup\/(?:setups\/[^/]+\/(?:sync|dns-setup)|clients\/[^/]+\/(?:dns-setup|dmarc))/.test(request.url()))providerRequests.push(request.url());if(request.method()==='POST'&&/\/clients\/[^/]+\/password\/reveal$/.test(request.url()))reveals.push(request.url());});await page.route('**/*',route=>route.request().url().startsWith(origin)?route.continue():route.abort());}
  const passwords=['local-QA-only-A!901','local-QA-only-B!902','local-QA-only-C!903','local-QA-conflict!904'];
  const clean=value=>{const text=typeof value==='string'?value:JSON.stringify(value);for(const password of passwords)assert.equal(text.includes(password),false,'normal state/HTML never exposes the password');if(value?.state)for(const client of value.state.clients){assert.equal('password'in client,false);assert.ok(Object.keys(client).every(key=>!/cipher|encrypted|secret/i.test(key)),'normal state must not expose stored encrypted credentials');}};
  const idle=page=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
  const data=async page=>{const model=await page.evaluate(()=>fetch('/api/gmail-setup').then(response=>response.json()));clean(model);return model;};
  const requestChange=(page,action,payload)=>page.evaluate(async({action,payload})=>{const board=await fetch('/api/gmail-setup').then(response=>response.json()),response=await fetch('/api/gmail-setup/change',{method:'POST',headers:{'Content-Type':'application/json','X-Domain-CSRF':document.querySelector('#gs-app').dataset.csrf},body:JSON.stringify({revision:board.revision,action,data:payload})});return {status:response.status,result:await response.json()};},{action,payload});
  const change=async(page,action,payload)=>{const response=await requestChange(page,action,payload);assert.equal(response.status,200,JSON.stringify(response.result));clean(response.result);return response.result;};
  const open=async(page,id)=>{const response=await page.goto(origin+`/gmail-set?passwordqa=${Date.now()}${id?'#client='+id:''}`);clean(await response.text());await idle(page);};
  const row=(model,id)=>model.state.clients.find(client=>client.id===id),input=page=>page.locator('#gs-client-password');
  const masked=async page=>{assert.equal(await input(page).getAttribute('type'),'password');assert.equal(await input(page).inputValue(),'');assert.equal(await input(page).getAttribute('autocomplete'),'new-password');};
  const passwordResponse=page=>page.waitForResponse(response=>response.request().method()==='POST'&&response.url().endsWith('/api/gmail-setup/change')&&response.request().postDataJSON().action==='client_password');
  const save=async(page,password)=>{if(password!==undefined)await input(page).fill(password);const saved=passwordResponse(page);await page.click('#gs-save-client-password');const response=await saved;assert.equal(response.status(),200);clean(await response.json());await idle(page);await masked(page);assert.equal(await input(page).getAttribute('placeholder'),'Password saved');};
  const show=async(page,password)=>{const response=page.waitForResponse(response=>response.request().method()==='POST'&&/\/clients\/[^/]+\/password\/reveal$/.test(response.url()));await page.click('#gs-show-client-password');const revealed=await response;assert.equal(revealed.status(),200);assert.equal(revealed.headers()['cache-control'],'no-store');assert.equal((await revealed.json()).password,password);await idle(page);assert.equal(await input(page).getAttribute('type'),'text');assert.equal(await input(page).inputValue(),password);};
  await open(kelton);const members=Object.fromEntries((await data(kelton)).state.members.map(member=>[member.name,member.id]));
  let model=await change(kelton,'client_save',{name:'Shared client password',section:'client',account_manager_id:members.Kelton,operator_id:members.Robyn,domain_target:2});const client=model.state.clients.at(-1);
  model=await change(kelton,'setup_bulk',{client_id:client.id,domains:'passwordqaone.com\npasswordqatwo.com',sender_name:'me',operator_id:members.Robyn});const job=model.state.setups[0];model=await change(kelton,'client_save',{name:'Password navigation neighbour',section:'client',account_manager_id:members.Kelton,operator_id:members.Robyn});const neighbour=model.state.clients.at(-1),baseline=structuredClone(model.summary);
  await open(kelton,client.id);await masked(kelton);assert.equal(row(await data(kelton),client.id).has_password,false);assert.equal(await input(kelton).getAttribute('placeholder'),'Add password');assert.equal(reveals.length,0);
  await save(kelton,passwords[0]);model=await data(kelton);assert.equal(row(model,client.id).has_password,true);assert.ok(row(model,client.id).password_updated_at);assert.deepEqual(model.summary,baseline,'saving a password does not change setup/device calculations');clean(await kelton.content());
  const empty=await requestChange(kelton,'client_password',{id:client.id,password:''});assert.equal(empty.status,400,'removing requires explicit null, not an empty Save');clean(empty.result);assert.equal(row(await data(kelton),client.id).has_password,true);
  await kelton.reload();await idle(kelton);await masked(kelton);clean(await kelton.content());assert.equal(reveals.length,0,'page load and board refresh never reveal automatically');await open(robyn,client.id);await masked(robyn);
  await show(kelton,passwords[0]);assert.equal(reveals.length,1);await kelton.click('#gs-show-client-password');await masked(kelton);assert.equal(reveals.length,1,'Hide does not request the saved password again');
  await show(kelton,passwords[0]);await input(kelton).fill(passwords[1]);await kelton.click('#gs-show-client-password');assert.equal(await input(kelton).getAttribute('type'),'password');assert.equal(await input(kelton).inputValue(),passwords[1],'Hide masks but preserves an edited replacement');await save(kelton);
  // A stale teammate cannot overwrite a newer password unknowingly. Their
  // unsaved replacement remains available until explicitly cleared/reviewed.
  await input(robyn).fill(passwords[3]);const conflict=passwordResponse(robyn);await robyn.click('#gs-save-client-password');assert.equal((await conflict).status(),409);await idle(robyn);assert.match(await robyn.textContent('#gs-message'),/password.*updated.*another|another.*password/i);assert.equal(await input(robyn).inputValue(),passwords[3]);await input(robyn).fill('');await show(robyn,passwords[1]);await robyn.reload();await idle(robyn);await masked(robyn);
  // Password drafts survive a regular board redraw, and a new release waits
  // for that draft to be saved before reloading the tool.
  await input(kelton).fill(passwords[2]);await kelton.selectOption(`[data-set-stage="${job.id}"]`,'Started');await idle(kelton);assert.equal(await input(kelton).inputValue(),passwords[2]);assert.equal(await input(kelton).getAttribute('type'),'password');
  let navigations=0;kelton.on('framenavigated',frame=>{if(frame===kelton.mainFrame())navigations++;});await kelton.evaluate(()=>document.activeElement?.blur());
  await kelton.route('**/api/gmail-setup/session',async route=>{const response=await route.fetch(),body=await response.json();await route.fulfill({response,json:{...body,app_version:'qa-release-with-password-draft'}});},{times:1});
  const checked=kelton.waitForResponse(response=>response.request().method()==='GET'&&response.url().endsWith('/api/gmail-setup/session'));await kelton.evaluate(()=>window.dispatchEvent(new PageTransitionEvent('pageshow',{persisted:true})));await (await checked).finished();await kelton.locator('#gs-update').waitFor({state:'visible'});assert.equal(await kelton.locator('#gs-apply-update').isDisabled(),true);assert.equal(navigations,0);assert.equal(await input(kelton).inputValue(),passwords[2]);
  const reloaded=kelton.waitForEvent('framenavigated',{predicate:frame=>frame===kelton.mainFrame()}),saved=passwordResponse(kelton);await kelton.click('#gs-save-client-password');clean(await (await saved).json());await reloaded;await kelton.waitForLoadState('domcontentloaded');await idle(kelton);await masked(kelton);clean(await kelton.content());assert.equal(await kelton.inputValue(`[data-set-stage="${job.id}"]`),'Started');
  await open(robyn,client.id);await show(robyn,passwords[2]);await robyn.click('#gs-show-client-password');await masked(robyn);
  await robyn.locator('#gs-client-info').screenshot({path:'work/gmail-client-password-masked-desktop.png',animations:'disabled'});await robyn.setViewportSize({width:390,height:844});assert.equal(await robyn.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false,'client password controls fit mobile');await robyn.locator('#gs-client-info').screenshot({path:'work/gmail-client-password-masked-mobile.png',animations:'disabled'});
  // Permission/workspace boundaries apply to explicit reveals as they do to
  // normal reads. A genuine missing-CSRF request is also rejected.
  const path=origin+`/api/gmail-setup/clients/${client.id}/password/reveal`;
  const csrf=await contexts[0].request.post(path,{data:{}});assert.equal(csrf.status(),403);clean(await csrf.json());
  const denied=await browser.newContext();await denied.addCookies([{name:'dashboard_session',value:cookie('Chris',{tools:['campaigns']}),url:origin,httpOnly:true,sameSite:'Lax'}]);const deniedResponse=await denied.request.post(path,{headers:{'X-Domain-CSRF':'chris-password-csrf'},data:{}});assert.equal(deniedResponse.status(),403);clean(await deniedResponse.json());
  const outside=await browser.newContext();await outside.addCookies([{name:'dashboard_session',value:cookie('Outside',{email:'outside@another-workspace.com'}),url:origin,httpOnly:true,sameSite:'Lax'}]);const outsideResponse=await outside.request.post(path,{headers:{'X-Domain-CSRF':'outside-password-csrf'},data:{}});assert.equal(outsideResponse.status(),404);clean(await outsideResponse.json());
  const removed=passwordResponse(robyn);await robyn.click('#gs-remove-client-password');const removal=await removed;assert.equal(removal.status(),200);assert.equal(removal.request().postDataJSON().data.password,null);clean(await removal.json());await idle(robyn);await masked(robyn);assert.equal(row(await data(robyn),client.id).has_password,false);assert.equal(await input(robyn).getAttribute('placeholder'),'Add password');assert.equal(await robyn.locator('#gs-show-client-password').isDisabled(),true);
  await kelton.reload();await idle(kelton);await masked(kelton);assert.equal(row(await data(kelton),client.id).has_password,false);const gone=await contexts[0].request.post(path,{headers:{'X-Domain-CSRF':'kelton-password-csrf'},data:{}});assert.equal(gone.status(),404);clean(await gone.json());clean(await kelton.content());
  // Browser Back/Forward can change the selected client while a request is
  // pending even though form controls are disabled. Late results must render
  // the current screen, never reattach the old client's password controls.
  const holdPassword=async()=>{let release,seen;const gate=new Promise(resolve=>release=resolve),captured=new Promise(resolve=>seen=resolve);await kelton.route('**/api/gmail-setup/change',async route=>{const response=await route.fetch();seen();await gate;await route.fulfill({response});},{times:1});return {captured,release:async()=>{release();await idle(kelton);}};};
  await input(kelton).fill(passwords[0]);let held=await holdPassword();await kelton.click('#gs-save-client-password');await held.captured;
  await kelton.evaluate(id=>{history.pushState(null,'',location.pathname+location.search+'#client='+id);dispatchEvent(new PopStateEvent('popstate'));},neighbour.id);assert.equal(await kelton.textContent('#gs-client-heading'),neighbour.name);await held.release();
  assert.equal(await kelton.textContent('#gs-client-heading'),neighbour.name);await masked(kelton);assert.equal(await input(kelton).getAttribute('placeholder'),'Add password');model=await data(kelton);assert.equal(row(model,client.id).has_password,true);assert.equal(row(model,neighbour.id).has_password,false);
  const onlyA=await contexts[1].request.post(path,{headers:{'X-Domain-CSRF':'robyn-password-csrf'},data:{}});assert.equal((await onlyA.json()).password,passwords[0]);const notB=await contexts[1].request.post(origin+`/api/gmail-setup/clients/${neighbour.id}/password/reveal`,{headers:{'X-Domain-CSRF':'robyn-password-csrf'},data:{}});assert.equal(notB.status(),404);
  await open(kelton,client.id);held=await holdPassword();await kelton.click('#gs-remove-client-password');await held.captured;await kelton.evaluate(()=>{history.pushState(null,'',location.pathname+location.search);dispatchEvent(new PopStateEvent('popstate'));});assert.equal(await kelton.locator('#gs-client-info').isVisible(),false);await held.release();assert.equal(await kelton.textContent('#gs-client-heading'),'Clients');assert.equal(await kelton.locator('#gs-client-info').isVisible(),false,'a late removal does not show the old client panel on the client list');
  model=await data(kelton);assert.equal(row(model,client.id).has_password,false);assert.equal(row(model,neighbour.id).has_password,false);
  assert.deepEqual(providerRequests,[]);assert.deepEqual(errors,[]);console.log('PASS: set/replace/remove shared client passwords, masked reload and explicit staff reveal, no normal API/HTML credential leaks, safe password conflicts, preserved drafts/stage edits/upgrades, tool/workspace/CSRF boundaries, mobile and zero provider requests.');
 }finally{if(browser)await browser.close();if(server&&server.exitCode===null){const exited=once(server,'exit');server.kill('SIGTERM');await Promise.race([exited,sleep(3000)]);}}
})().catch(error=>{console.error(error);process.exit(1)});
