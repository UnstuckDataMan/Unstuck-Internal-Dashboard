// Master-domain UI regression with genuine signed SDR users and fake storage.
// The private Python child blocks every external provider request.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
const {spawn,execFileSync}=require('node:child_process'),{once}=require('node:events');
const origin='http://127.0.0.1:8768',sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms));
const harness=`
import importlib,os,requests,uvicorn
os.environ['SUPABASE_URL']='http://supabase.invalid'
os.environ['SUPABASE_SERVICE_ROLE_KEY']=''
def forbidden(*args,**kwargs):raise AssertionError('External provider call forbidden in master-domain QA')
requests.sessions.Session.request=forbidden
from tests.gmail_setup_preview import app
from app import auth
from app.utils.gmail_setup import store
importlib.reload(auth)
os.environ['AUTH_DISABLED']='0'
store.team_directory=lambda user:[{'name':name,'email':name.lower()+'@unstuck-agency.com'} for name in ['Dylan','Kelton','Robyn','Chris']]
uvicorn.run(app,host='127.0.0.1',port=8768,log_level='warning',lifespan='off')
`;
const cookie=name=>execFileSync('.venv/bin/python',['-c',"import base64,json,sys;from itsdangerous import TimestampSigner;print(TimestampSigner('gmail-local-qa').sign(base64.b64encode(json.dumps(json.loads(sys.argv[1])).encode())).decode())",JSON.stringify({user:{email:name.toLowerCase()+'@unstuck-agency.com',name,role:'sdr',tools:['gmail_set']},domain_csrf:name.toLowerCase()+'-master-csrf'})],{encoding:'utf8'}).trim();
(async()=>{
 let server,browser,log='';
 try{
  server=spawn('.venv/bin/python',['-c',harness],{stdio:['ignore','pipe','pipe']});server.stdout.on('data',data=>log+=data);server.stderr.on('data',data=>log+=data);
  let ready=false;for(let i=0;i<100;i++){if(server.exitCode!==null)throw Error(log);try{ready=(await fetch(origin+'/healthz')).ok;}catch{}if(ready)break;await sleep(100);}assert.equal(ready,true);
  browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});const contexts=await Promise.all([browser.newContext({viewport:{width:1440,height:1050}}),browser.newContext({viewport:{width:1440,height:1050}})]);
  for(const [index,name] of ['Kelton','Robyn'].entries())await contexts[index].addCookies([{name:'dashboard_session',value:cookie(name),url:origin,httpOnly:true,sameSite:'Lax'}]);
  const [kelton,robyn]=await Promise.all(contexts.map(context=>context.newPage())),errors=[],providerRequests=[];
  for(const page of [kelton,robyn]){page.on('pageerror',error=>errors.push(error.message));page.on('dialog',dialog=>dialog.accept());page.on('request',request=>{if(/\/api\/gmail-setup\/(?:setups\/[^/]+\/dns-setup|clients\/[^/]+\/(?:dns-setup|dmarc))/.test(request.url()))providerRequests.push(request.url());});await page.route('**/*',route=>route.request().url().startsWith(origin)?route.continue():route.abort());}
  const idle=page=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
  const data=page=>page.evaluate(()=>fetch('/api/gmail-setup').then(response=>response.json()));
  const requestChange=(page,action,payload)=>page.evaluate(async({action,payload})=>{const board=await fetch('/api/gmail-setup').then(response=>response.json()),response=await fetch('/api/gmail-setup/change',{method:'POST',headers:{'Content-Type':'application/json','X-Domain-CSRF':document.querySelector('#gs-app').dataset.csrf},body:JSON.stringify({revision:board.revision,action,data:payload})});return {status:response.status,result:await response.json()};},{action,payload});
  const change=async(page,action,payload)=>{const response=await requestChange(page,action,payload);assert.equal(response.status,200,JSON.stringify(response.result));return response.result;};
  const open=async(page,id)=>{await page.goto(origin+`/gmail-set?masterqa=${Date.now()}${id?'#client='+id:''}`);await idle(page);};
  const card=(page,id)=>page.locator(`[data-setup="${id}"]`),selector=(page,id)=>page.locator(`[data-set-master="${id}"]`);
  const select=async(page,id,master)=>{await selector(page,id).selectOption(master||'');await idle(page);};
  const stage=async(id,value)=>{await kelton.selectOption(`[data-set-stage="${id}"]`,value);await idle(kelton);};
  const stored=(model,id)=>model.state.setups.find(job=>job.id===id);
  const background=locator=>locator.evaluate(element=>{const css=getComputedStyle(element);return {color:css.backgroundColor,image:css.backgroundImage};});
  const buttonStyle=locator=>locator.evaluate(element=>{const css=getComputedStyle(element);return {background:css.backgroundColor,color:css.color,border:css.borderColor};});
  const group=async(page,id,root)=>{const node=card(page,id),master=await node.getAttribute('data-master-root');if(!root){assert.ok(!master);assert.doesNotMatch(await node.getAttribute('class'),/\bgs-master-[0-4]\b/);return null;}assert.equal(master,root);const slot=await node.getAttribute('data-master-color');assert.match(slot,/^[0-4]$/);assert.match(await node.getAttribute('class'),new RegExp('\\bgs-master-'+slot+'\\b'));return slot;};
  await open(kelton);const members=Object.fromEntries((await data(kelton)).state.members.map(member=>[member.name,member.id]));
  let model=await change(kelton,'client_save',{name:'Master domain groups',section:'client',account_manager_id:members.Kelton,operator_id:members.Robyn}),client=model.state.clients.at(-1);
  model=await change(kelton,'setup_bulk',{client_id:client.id,domains:Array.from({length:12},(_,index)=>`masterqa${index}.com`).join('\n'),sender_name:'me',operator_id:members.Robyn});const jobs=model.state.setups.filter(job=>job.client_id===client.id),roots=jobs.slice(0,5);
  model=await change(kelton,'client_save',{name:'Separate master client',section:'client',account_manager_id:members.Kelton,operator_id:members.Robyn});const otherClient=model.state.clients.at(-1);
  model=await change(kelton,'setup_bulk',{client_id:otherClient.id,domains:'masterotherone.com\nmasterothertwo.com',sender_name:'me',operator_id:members.Robyn});const others=model.state.setups.filter(job=>job.client_id===otherClient.id),baseline=structuredClone(model.summary);
  await Promise.all([open(kelton,client.id),open(robyn,client.id)]);
  const neutral=await background(card(kelton,roots[0].id)),normalButton=await buttonStyle(card(kelton,roots[0].id).locator('[data-remove-setup]'));
  for(const root of roots)await select(kelton,root.id,root.id);
  const slots=await Promise.all(roots.map(root=>group(kelton,root.id,root.id)));assert.equal(new Set(slots).size,5,'each of the five roots has its own strip colour');
  for(const root of roots){assert.equal(await card(kelton,root.id).locator('svg[aria-label="Master domain"]').count(),1);assert.deepEqual(await background(card(kelton,root.id)),neutral);assert.deepEqual(await buttonStyle(card(kelton,root.id).locator('[data-remove-setup]')),normalButton);}
  assert.deepEqual((await data(kelton)).summary,baseline,'master organisation does not change demand or device calculations');
  const sixthOption=selector(kelton,jobs[5].id).locator(`option[value="${jobs[5].id}"]`);assert.equal(await sixthOption.evaluate(option=>option.disabled),true,'a sixth master cannot be selected');
  const tooMany=await requestChange(kelton,'setup_master',{id:jobs[5].id,master_domain_id:jobs[5].id});assert.ok([400,409].includes(tooMany.status));assert.match(JSON.stringify(tooMany.result),/5|five/i);assert.equal(stored(await data(kelton),jobs[5].id).master_domain_id??null,null);
  // Another client's own five-master allowance is separate; cross-client
  // links are rejected even if crafted outside the dropdown.
  await open(robyn,otherClient.id);await select(robyn,others[0].id,others[0].id);assert.equal(await selector(robyn,others[1].id).locator(`option[value="${roots[0].id}"]`).count(),0);
  const crossed=await requestChange(robyn,'setup_master',{id:others[1].id,master_domain_id:roots[0].id});assert.ok([400,409].includes(crossed.status));assert.match(JSON.stringify(crossed.result),/client/i);assert.equal(stored(await data(robyn),others[1].id).master_domain_id??null,null);
  // Link, change and undo a child. Only the root receives a crown.
  await open(kelton,client.id);await select(kelton,jobs[5].id,roots[0].id);assert.equal(await group(kelton,jobs[5].id,roots[0].id),slots[0]);assert.equal(await card(kelton,jobs[5].id).locator('svg[aria-label="Master domain"]').count(),0);
  await select(kelton,jobs[5].id,roots[1].id);assert.equal(await group(kelton,jobs[5].id,roots[1].id),slots[1]);await select(kelton,jobs[5].id,null);await group(kelton,jobs[5].id,null);
  // Stage backgrounds remain independent of master strips, including closed
  // records: completion/burning does not prevent assigning or undoing a group.
  await stage(roots[0].id,'Started');await stage(roots[1].id,'Setup complete');await stage(roots[2].id,'Burned');
  await stage(jobs[6].id,'Setup complete');const completeShade=await background(card(kelton,jobs[6].id));assert.equal(await selector(kelton,jobs[6].id).isEnabled(),true);await select(kelton,jobs[6].id,roots[1].id);assert.deepEqual(await background(card(kelton,jobs[6].id)),completeShade);assert.match(await card(kelton,jobs[6].id).getAttribute('class'),/gs-domain-complete/);await select(kelton,jobs[6].id,null);await group(kelton,jobs[6].id,null);
  await stage(jobs[7].id,'Burned');const burnedShade=await background(card(kelton,jobs[7].id));assert.equal(await selector(kelton,jobs[7].id).isEnabled(),true);await select(kelton,jobs[7].id,roots[2].id);assert.deepEqual(await background(card(kelton,jobs[7].id)),burnedShade);assert.match(await card(kelton,jobs[7].id).getAttribute('class'),/gs-domain-danger/);assert.equal(await card(kelton,roots[2].id).locator('svg[aria-label="Master domain"]').count(),1);
  assert.notDeepEqual(completeShade,burnedShade);assert.notDeepEqual(completeShade,neutral);assert.deepEqual(await buttonStyle(card(kelton,jobs[7].id).locator('[data-remove-setup]')),normalButton);
  // Demoting a root ungroups its children; other roots retain their colours.
  await select(kelton,jobs[5].id,roots[0].id);await select(kelton,jobs[6].id,roots[0].id);const startedShade=await background(card(kelton,roots[0].id));await select(kelton,roots[0].id,null);
  for(const job of [roots[0],jobs[5],jobs[6]]){await group(kelton,job.id,null);assert.equal(stored(await data(kelton),job.id).master_domain_id??null,null);}assert.deepEqual(await background(card(kelton,roots[0].id)),startedShade);
  for(let index=1;index<roots.length;index++)assert.equal(await group(kelton,roots[index].id,roots[index].id),slots[index]);
  await select(kelton,jobs[8].id,jobs[8].id);assert.equal(await group(kelton,jobs[8].id,jobs[8].id),slots[0],'the freed colour is available to a new master');
  // Removing a root cleans links on every child without removing those rows.
  await select(kelton,jobs[9].id,roots[3].id);await select(kelton,jobs[10].id,roots[3].id);await card(kelton,roots[3].id).locator('[data-remove-setup]').click();await kelton.waitForFunction(()=>document.querySelector('#gs-dialog').open);await kelton.click('#gs-dialog-save');await kelton.waitForFunction(()=>!document.querySelector('#gs-dialog').open);await idle(kelton);
  for(const job of [jobs[9],jobs[10]]){await group(kelton,job.id,null);assert.equal(stored(await data(kelton),job.id).master_domain_id??null,null);}assert.equal(stored(await data(kelton),roots[3].id).archived,true);assert.equal(await selector(kelton,jobs[9].id).locator(`option[value="${roots[3].id}"]`).count(),0);
  await select(kelton,jobs[5].id,jobs[8].id);await stage(jobs[8].id,'Started');await stage(jobs[5].id,'Started');
  // A separate signed-in teammate sees the same links after refresh/reload.
  await open(robyn,client.id);for(const page of [kelton,robyn]){assert.equal(await group(page,jobs[5].id,jobs[8].id),slots[0]);assert.equal(await group(page,jobs[8].id,jobs[8].id),slots[0]);assert.equal(await selector(page,jobs[5].id).inputValue(),jobs[8].id);}
  await Promise.all([kelton.reload(),robyn.reload()]);await Promise.all([idle(kelton),idle(robyn)]);assert.equal(await group(robyn,jobs[5].id,jobs[8].id),slots[0]);assert.equal(await card(robyn,jobs[8].id).locator('svg[aria-label="Master domain"]').count(),1);assert.equal(await card(robyn,jobs[5].id).locator('svg[aria-label="Master domain"]').count(),0);
  await card(kelton,jobs[8].id).screenshot({path:'work/gmail-master-root-started.png',animations:'disabled'});await card(kelton,jobs[5].id).screenshot({path:'work/gmail-master-child-started.png',animations:'disabled'});await card(kelton,roots[1].id).screenshot({path:'work/gmail-master-root-complete.png',animations:'disabled'});await card(kelton,roots[2].id).screenshot({path:'work/gmail-master-root-burned.png',animations:'disabled'});
  await kelton.setViewportSize({width:390,height:844});assert.equal(await kelton.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false,'master controls fit mobile');await card(kelton,jobs[8].id).screenshot({path:'work/gmail-master-mobile.png',animations:'disabled'});
  // Keep a compact visual overview: one root, two children and an ungrouped
  // neighbour make the left-strip relationship easy to inspect together.
  model=await change(kelton,'client_save',{name:'Master grouping visual overview',section:'client',account_manager_id:members.Kelton,operator_id:members.Robyn});const visualClient=model.state.clients.at(-1);
  model=await change(kelton,'setup_bulk',{client_id:visualClient.id,domains:'groupmasterqa.com\ngroupfirstqa.com\ngroupsecondqa.com\nungroupedqa.com',sender_name:'me',operator_id:members.Robyn});const visual=model.state.setups.filter(job=>job.client_id===visualClient.id);
  await kelton.setViewportSize({width:1440,height:2000});await open(kelton,visualClient.id);await select(kelton,visual[0].id,visual[0].id);await select(kelton,visual[1].id,visual[0].id);await select(kelton,visual[2].id,visual[0].id);await stage(visual[0].id,'Started');await stage(visual[1].id,'Started');
  await group(kelton,visual[3].id,null);await kelton.locator('#gs-setup-list').evaluate(element=>element.scrollIntoView({block:'start'}));await kelton.screenshot({path:'work/gmail-master-overview-desktop.png',animations:'disabled'});
  assert.deepEqual(providerRequests,[]);assert.deepEqual(errors,[]);console.log('PASS: five-master limit, independent client groups, crown/strip inheritance, child relink/undo, demotion/delete cleanup, stable colours, complete/burned controls, preserved stage/button styles, signed teammate persistence, mobile and zero registrar requests.');
 }finally{if(browser)await browser.close();if(server&&server.exitCode===null){const exited=once(server,'exit');server.kill('SIGTERM');await Promise.race([exited,sleep(3000)]);}}
})().catch(error=>{console.error(error);process.exit(1)});
