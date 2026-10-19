// Local-only shared-device regression. Separate browser contexts share only
// the synthetic preview server's board; all external requests are blocked.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
(async()=>{
 const browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});
 const contexts=await Promise.all([browser.newContext(),browser.newContext()]),[one,two]=await Promise.all(contexts.map(c=>c.newPage()));
 const errors=[],syncRequests=[];
 for(const page of [one,two]){
  page.on('pageerror',e=>errors.push(e.message));
  page.on('request',r=>{if(r.method()==='POST'&&/\/api\/gmail-setup\/setups\/[^/]+\/sync$/.test(r.url()))syncRequests.push(r.url());});
  await page.route('**/*',r=>r.request().url().startsWith('http://127.0.0.1:8766')?r.continue():r.abort());
 }
 const idle=page=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
 const within=async promise=>{let timer;try{return await Promise.race([promise,new Promise((_,reject)=>timer=setTimeout(()=>reject(Error('Background board read did not start.')),10000))]);}finally{clearTimeout(timer);}};
 const data=page=>page.evaluate(()=>fetch('/api/gmail-setup').then(r=>r.json()));
 const change=(page,action,payload)=>page.evaluate(async({action,payload})=>{
  const board=await fetch('/api/gmail-setup').then(r=>r.json());
  const response=await fetch('/api/gmail-setup/change',{method:'POST',headers:{'Content-Type':'application/json','X-Domain-CSRF':document.querySelector('#gs-app').dataset.csrf},body:JSON.stringify({revision:board.revision,action,data:payload})});
  const result=await response.json();if(!response.ok)throw Error(JSON.stringify(result));return result;
 },{action,payload});
 const open=async(page,id)=>{await page.goto(`http://127.0.0.1:8766/gmail-set?sharedqa=${Date.now()}#client=${id}`);await idle(page);await page.waitForSelector('[data-setup]');};
 const stage=async(page,id,value)=>{await page.selectOption(`[data-set-stage="${id}"]`,value);await idle(page);};
 const select=(page,id)=>page.locator(`[data-set-stage="${id}"]`);
 const trackChanges=page=>{const changes=[];const listener=response=>{if(response.request().method()==='POST'&&response.url().endsWith('/api/gmail-setup/change'))changes.push({status:response.status(),body:response.request().postDataJSON()});};page.on('response',listener);return {changes,stop:()=>page.off('response',listener)};};
 const isBoard=response=>response.request().method()==='GET'&&new URL(response.url()).pathname==='/api/gmail-setup';
 const poll=async page=>{const response=page.waitForResponse(isBoard);await page.evaluate(()=>{document.activeElement?.blur();document.dispatchEvent(new Event('visibilitychange'));});await (await response).finished();};
 await one.goto('http://127.0.0.1:8766/gmail-set');await idle(one);
 const tag=String(Date.now()),owner=(await data(one)).state.members.find(m=>m.active).id;
 const created=await change(one,'client_save',{name:`Shared devices ${tag}`,section:'client',account_manager_id:owner,operator_id:owner}),client=created.state.clients.at(-1);
 await change(one,'setup_bulk',{client_id:client.id,domains:['first','second','delayed'].map(x=>`shared${x}${tag}.com`).join('\n'),sender_name:'sender',operator_id:owner});
 let model=await data(one);const jobs=model.state.setups.filter(j=>j.client_id===client.id),[first,second,delayed]=jobs;
 await Promise.all([open(one,client.id),open(two,client.id)]);
 // Device two has a stale workspace revision after device one saves. Its
 // unrelated domain edit should rebase, retaining both people's changes.
 await stage(one,first.id,'Started');const unrelated=trackChanges(two);await stage(two,second.id,'Started');unrelated.stop();
 assert.equal(unrelated.changes[0]?.status,409,'exercise the stale-revision conflict');assert.equal(unrelated.changes.at(-1)?.status,200,'an unrelated edit is retried safely');assert.ok(unrelated.changes.length<=3,'retries are bounded');
 model=await data(one);assert.equal(model.state.setups.find(j=>j.id===first.id).stage,'Started');assert.equal(model.state.setups.find(j=>j.id===second.id).stage,'Started');
 assert.equal(await select(two,first.id).inputValue(),'Started');assert.equal(await select(two,second.id).inputValue(),'Started');
 await poll(one);await one.waitForFunction(id=>document.querySelector(`[data-set-stage="${id}"]`)?.value==='Started',second.id);
 await one.reload();await idle(one);assert.equal(await select(one,first.id).inputValue(),'Started');assert.equal(await select(one,second.id).inputValue(),'Started');
 // Both devices now saw Started. A newer change to this same domain must win;
 // the stale device must report the conflict rather than resetting its stage.
 await stage(one,first.id,'Purchase made');const same=trackChanges(two);await stage(two,first.id,'Not started');same.stop();
 assert.deepEqual(same.changes.map(c=>c.status),[409],'a changed target stage is never silently overwritten');
 assert.equal((await data(two)).state.setups.find(j=>j.id===first.id).stage,'Purchase made');assert.equal(await select(two,first.id).inputValue(),'Purchase made');
 assert.match(await two.textContent('#gs-message'),/changed|updated|another|teammate|conflict/i);assert.equal(await two.locator('#gs-message').evaluate(el=>el.classList.contains('error')),true);
 assert.equal(await select(two,second.id).inputValue(),'Started','another device changing a different domain preserves Started');
 // A background read can begin before the user focuses a dropdown. Even if
 // another domain makes that response's revision newer than the local board,
 // it must not erase a choice the user has made since the read began.
 await poll(one);await change(two,'setup_notes',{id:second.id,notes:'Another device edited this domain while the first device was viewing the board.'});
 let release,seen;const gate=new Promise(r=>release=r),captured=new Promise(r=>seen=r);
 const handler=async route=>{const response=await route.fetch();seen();await gate;await route.fulfill({response});};
 await one.route('**/api/gmail-setup',handler,{times:1});
 const background=one.waitForResponse(isBoard);await one.evaluate(()=>{document.activeElement?.blur();document.dispatchEvent(new Event('visibilitychange'));});
 await within(captured);const original=await select(one,delayed.id).elementHandle();await select(one,delayed.id).focus();await select(one,delayed.id).evaluate(el=>{el.value='Started';});
 release();await (await background).finished();await one.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
 assert.equal(await original.evaluate(el=>el.isConnected),true,'a delayed background response must not replace the focused dropdown');assert.equal(await select(one,delayed.id).inputValue(),'Started','the selected value survives the delayed read');
 await select(one,delayed.id).dispatchEvent('change');await idle(one);assert.equal((await data(one)).state.setups.find(j=>j.id===delayed.id).stage,'Started');
 await poll(two);await two.waitForFunction(id=>document.querySelector(`[data-set-stage="${id}"]`)?.value==='Started',delayed.id);
 await Promise.all([one.reload(),two.reload()]);await Promise.all([idle(one),idle(two)]);
 for(const page of [one,two]){assert.equal(await select(page,first.id).inputValue(),'Purchase made');assert.equal(await select(page,second.id).inputValue(),'Started');assert.equal(await select(page,delayed.id).inputValue(),'Started');}
 assert.deepEqual(syncRequests,[]);assert.deepEqual(errors,[]);await browser.close();console.log('PASS: two independent devices safely rebase unrelated changes, reject conflicting stage overwrites, retain Started across another user/poll/reload, protect an active dropdown from delayed reads, and keep sheet updates manual.');
})().catch(e=>{console.error(e);process.exit(1)});
