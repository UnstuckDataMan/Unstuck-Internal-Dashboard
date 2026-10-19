// Authenticated homepage/cache regression in a private synthetic server.
// The app uses real signed SDR sessions; every provider is isolated.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
const {spawn,execFileSync}=require('node:child_process'),{once}=require('node:events'),{createHash}=require('node:crypto');
const origin='http://127.0.0.1:8768',sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms));
const harness=`
import importlib,os,requests,uvicorn
os.environ['SUPABASE_URL']='http://supabase.invalid'
os.environ['SUPABASE_SERVICE_ROLE_KEY']=''
os.environ['RENDER_GIT_COMMIT']='ffffffffffffffffffffffffffffffffffffffff'
def forbidden(*args,**kwargs):raise AssertionError('External provider call forbidden in homepage QA')
requests.sessions.Session.request=forbidden
from tests.gmail_setup_preview import app
from app import auth
from app.utils.gmail_setup import store
importlib.reload(auth)
os.environ['AUTH_DISABLED']='0'
store.team_directory=lambda user:[{'name':name,'email':name.lower()+'@unstuck-agency.com'} for name in ['Dylan','Kelton','Robyn','Chris']]
uvicorn.run(app,host='127.0.0.1',port=8768,log_level='warning',lifespan='off')
`;
const cookie=name=>execFileSync('.venv/bin/python',['-c',"import base64,json,sys;from itsdangerous import TimestampSigner;print(TimestampSigner('gmail-local-qa').sign(base64.b64encode(json.dumps(json.loads(sys.argv[1])).encode())).decode())",JSON.stringify({user:{email:name.toLowerCase()+'@unstuck-agency.com',name,role:'sdr',tools:['gmail_set']},domain_csrf:name.toLowerCase()+'-home-csrf'})],{encoding:'utf8'}).trim();
(async()=>{
 let server,browser,log='';
 try{
  server=spawn('.venv/bin/python',['-c',harness],{stdio:['ignore','pipe','pipe']});server.stdout.on('data',data=>log+=data);server.stderr.on('data',data=>log+=data);
  let ready=false;for(let i=0;i<100;i++){if(server.exitCode!==null)throw Error(log);try{ready=(await fetch(origin+'/healthz')).ok;}catch{}if(ready)break;await sleep(100);}assert.equal(ready,true);
  browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});
  const contexts=await Promise.all([browser.newContext({viewport:{width:1440,height:1050},reducedMotion:'reduce'}),browser.newContext({viewport:{width:1440,height:1050},reducedMotion:'reduce'})]);
  await contexts[0].addCookies([{name:'dashboard_session',value:cookie('Kelton'),url:origin,httpOnly:true,sameSite:'Lax'}]);await contexts[1].addCookies([{name:'dashboard_session',value:cookie('Robyn'),url:origin,httpOnly:true,sameSite:'Lax'}]);
  const kelton=await contexts[0].newPage(),robyn=await contexts[1].newPage(),errors=[],syncRequests=[];
  const protect=async page=>{page.on('pageerror',error=>errors.push(error.message));page.on('request',r=>{if(r.method()==='POST'&&/\/setups\/[^/]+\/sync$/.test(r.url()))syncRequests.push(r.url());});await page.route('**/*',route=>route.request().url().startsWith(origin)?route.continue():route.abort());};
  await Promise.all([protect(kelton),protect(robyn)]);
  await kelton.addInitScript(()=>{const original=window.setInterval;window.__homeQAIntervals=[];window.setInterval=(callback,delay,...args)=>{window.__homeQAIntervals.push({delay,run:()=>callback(...args)});return original(callback,delay,...args);};});
  const stock=(page,text)=>page.waitForFunction(expected=>document.querySelector('#gs-home-stock')?.textContent===expected,text);
  const summaryResponse=page=>page.waitForResponse(r=>r.request().method()==='GET'&&new URL(r.url()).pathname==='/api/gmail-setup/summary');
  const signal=async(page,event)=>{const response=summaryResponse(page);await page.evaluate(event=>{if(event==='pageshow')window.dispatchEvent(new PageTransitionEvent('pageshow',{persisted:true}));else document.dispatchEvent(new Event('visibilitychange'));},event);await (await response).finished();};
  const idle=page=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
  const data=page=>page.evaluate(()=>fetch('/api/gmail-setup').then(r=>r.json()));
  const change=(page,action,payload)=>page.evaluate(async({action,payload})=>{const board=await fetch('/api/gmail-setup').then(r=>r.json());const response=await fetch('/api/gmail-setup/change',{method:'POST',headers:{'Content-Type':'application/json','X-Domain-CSRF':document.querySelector('#gs-app').dataset.csrf},body:JSON.stringify({revision:board.revision,action,data:payload})});const result=await response.json();if(!response.ok)throw Error(JSON.stringify(result));return result;},{action,payload});
  const homeResponse=await kelton.goto(origin+'/');assert.equal(homeResponse.headers()['cache-control'],'no-store');await stock(kelton,'Order 5 more phones');
  const script=await kelton.locator('script[src*="gmail-setup-home.js"]').getAttribute('src');assert.match(script,/^\/static\/gmail-setup-home\.js\?v=[a-f0-9]{12}$/);
  const asset=await kelton.request.get(origin+script);assert.equal(asset.headers()['cache-control'],'no-store');assert.equal(new URL(script,origin).searchParams.get('v'),createHash('sha256').update(await asset.body()).digest('hex').slice(0,12),'home asset version follows its actual contents');
  assert.equal(await kelton.evaluate(()=>window.__homeQAIntervals.some(timer=>timer.delay===30000)),true,'the Operations figure polls every 30 seconds');
  const clicked=summaryResponse(kelton);await kelton.click('#btn-operations');await (await clicked).finished();await stock(kelton,'Order 5 more phones');
  await robyn.goto(origin+'/gmail-set');await idle(robyn);const members=Object.fromEntries((await data(robyn)).state.members.map(member=>[member.name,member.id]));
  let model=await change(robyn,'client_save',{name:'Homepage live figures',section:'client',account_manager_id:members.Kelton,operator_id:members.Robyn,domain_target:10,sender_full_name:'Richard Home',website_url:'home-client.example.com',account_manager_email:'kelton@unstuck-agency.com'}),client=model.state.clients.at(-1);
  model=await change(robyn,'setup_bulk',{client_id:client.id,domains:'homepageone.com\nhomepagetwo.com',sender_name:'richard',operator_id:members.Robyn});const first=model.state.setups[0];
  // The other user's new work changes demand. A restored homepage must read
  // current figures rather than showing the value preserved in browser memory.
  assert.equal(await kelton.textContent('#gs-home-stock'),'Order 5 more phones');await signal(kelton,'pageshow');await stock(kelton,'Order 10 more phones');
  await change(robyn,'devices_add',{member_id:members.Robyn,phones:1});assert.equal(await kelton.textContent('#gs-home-stock'),'Order 10 more phones');await signal(kelton,'visibilitychange');await stock(kelton,'Order 9 more phones');
  await change(robyn,'devices_add',{member_id:members.Robyn,phones:1});const periodic=summaryResponse(kelton);await kelton.evaluate(()=>window.__homeQAIntervals.find(timer=>timer.delay===30000).run());await (await periodic).finished();await stock(kelton,'Order 8 more phones');
  // A restore event can arrive during an earlier read. Coalesce those events
  // into a fresh follow-up, rather than losing them or issuing parallel reads.
  let release,seen,heldRequest,summaryReads=0;const gate=new Promise(resolve=>release=resolve),captured=new Promise(resolve=>seen=resolve);
  const countReads=request=>{if(request.method()==='GET'&&new URL(request.url()).pathname==='/api/gmail-setup/summary')summaryReads++;};kelton.on('request',countReads);
  await kelton.route('**/api/gmail-setup/summary',async route=>{heldRequest=route.request();const response=await route.fetch();seen();await gate;await route.fulfill({response});},{times:1});
  const olderRead=summaryResponse(kelton);await kelton.evaluate(()=>document.dispatchEvent(new Event('visibilitychange')));await captured;
  await change(robyn,'devices_add',{member_id:members.Robyn,phones:1});await kelton.evaluate(()=>{for(let i=0;i<3;i++)window.dispatchEvent(new PageTransitionEvent('pageshow',{persisted:true}));});assert.equal(summaryReads,1,'restore events do not create parallel summary requests');
  const followup=kelton.waitForResponse(response=>response.request().method()==='GET'&&new URL(response.url()).pathname==='/api/gmail-setup/summary'&&response.request()!==heldRequest);release();await (await olderRead).finished();await (await followup).finished();await stock(kelton,'Order 7 more phones');assert.equal(summaryReads,2,'a single fresh read follows the held stale response');kelton.off('request',countReads);
  // Failed reads must not leave a previous number looking current.
  await kelton.route('**/api/gmail-setup/summary',route=>route.fulfill({status:503,json:{detail:'Synthetic unavailable summary.'}}),{times:1});await signal(kelton,'pageshow');await stock(kelton,'Device needs unavailable');
  await signal(kelton,'visibilitychange');await stock(kelton,'Order 7 more phones');
  const robynHome=await contexts[1].newPage();await protect(robynHome);const otherHome=await robynHome.goto(origin+'/');assert.equal(otherHome.headers()['cache-control'],'no-store');await stock(robynHome,'Order 7 more phones');assert.equal(await robynHome.locator('script[src*="gmail-setup-home.js"]').getAttribute('src'),script,'both SDR accounts receive the same current homepage code');
  // The historical tool URL and current URL both serve the current layout,
  // exact release ID and uncached assets to ordinary authenticated users.
  const versions=[];for(const path of ['/gmail-set','/domain-cluster-builder']){const response=await robyn.request.get(origin+path),html=await response.text();assert.equal(response.status(),200);assert.equal(response.headers()['cache-control'],'no-store');versions.push({app:html.match(/data-app-version="([^"]+)"/)[1],asset:html.match(/src="([^\"]*gmail-setup\.js\?v=[a-f0-9]{12})"/)[1]});assert.match(html,/id="gs-client-info"/);}assert.deepEqual(versions[0],versions[1]);assert.equal(versions[0].app,'ffffffffffffffffffffffffffffffffffffffff');
  await robyn.goto(origin+`/gmail-set?homeqa=restore#client=${client.id}`);await idle(robyn);await robyn.fill(`[data-set-notes="${first.id}"]`,'Keep this unsaved note when restoring the old tool tab.');await robyn.evaluate(()=>document.activeElement?.blur());
  let navigations=0;robyn.on('framenavigated',frame=>{if(frame===robyn.mainFrame())navigations++;});
  await robyn.route('**/api/gmail-setup/session',async route=>{const response=await route.fetch(),body=await response.json();await route.fulfill({response,json:{...body,app_version:'qa-new-release-after-back'}});},{times:1});
  const checked=robyn.waitForResponse(r=>r.request().method()==='GET'&&r.url().endsWith('/api/gmail-setup/session'));await robyn.evaluate(()=>window.dispatchEvent(new PageTransitionEvent('pageshow',{persisted:true})));await (await checked).finished();await robyn.locator('#gs-update').waitFor({state:'visible'});
  assert.equal(navigations,0,'restoring an older tool tab preserves unsaved notes');assert.equal(await robyn.locator('#gs-apply-update').isDisabled(),true);assert.equal(await robyn.inputValue(`[data-set-notes="${first.id}"]`),'Keep this unsaved note when restoring the old tool tab.');
  const reloaded=robyn.waitForEvent('framenavigated',{predicate:frame=>frame===robyn.mainFrame()});await robyn.click(`[data-save-notes="${first.id}"]`);await reloaded;await robyn.waitForLoadState('domcontentloaded');await idle(robyn);assert.equal(await robyn.inputValue(`[data-set-notes="${first.id}"]`),'Keep this unsaved note when restoring the old tool tab.');
  assert.deepEqual(syncRequests,[]);assert.deepEqual(errors,[]);console.log('PASS: signed SDR homepage parity, no-store HTML and content-versioned JS, 30-second/current-on-return figures, queued restores after an in-flight read, stale-number clearing and recovery, current/legacy tool URL parity, and restored-tab upgrade protection for unsaved notes.');
 }finally{if(browser)await browser.close();if(server&&server.exitCode===null){const exited=once(server,'exit');server.kill('SIGTERM');await Promise.race([exited,sleep(3000)]);}}
})().catch(error=>{console.error(error);process.exit(1)});
