// Synthetic UI: stage shading belongs to the card; button styles stay stable.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
(async()=>{
 const browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});
 const page=await browser.newPage({viewport:{width:1440,height:1050}}),errors=[];
 page.on('pageerror',e=>errors.push(e.message));
 await page.route('**/*',r=>r.request().url().startsWith('http://127.0.0.1:8766')?r.continue():r.abort());
 const idle=()=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
 const data=()=>page.evaluate(()=>fetch('/api/gmail-setup').then(r=>r.json()));
 const stage=async(id,value)=>{await page.selectOption(`[data-set-stage="${id}"]`,value);await idle();};
 const style=async locator=>locator.evaluate(el=>{const s=getComputedStyle(el);return {background:s.backgroundColor,color:s.color,border:s.borderColor};});
 await page.goto('http://127.0.0.1:8766/gmail-set');await idle();const baseline=await data(),owner=baseline.state.members.find(m=>m.name==='Dylan').id;
 await page.click('#gs-add-setup');await page.fill('#gs-client-name','Stage colours');await page.selectOption('#gs-client-manager',owner);await page.selectOption('#gs-client-operator',owner);await page.click('#gs-dialog-save');await page.waitForFunction(()=>!document.querySelector('#gs-dialog').open);await idle();
 await page.fill('#gs-bulk-sender','me');await page.fill('#gs-bulk-domains','startedcolour.com\ncompletecolour.com\nburnedcolour.com\nqrcolour.com\nneutralcolour.com');await page.locator('#gs-bulk-panel button[type="submit"]').click();await idle();
 let model=await data();const jobs=model.state.setups.filter(j=>j.client_name==='Stage colours'),[started,complete,burned,qr,neutral]=jobs;
 const card=id=>page.locator(`[data-setup="${id}"]`),remove=id=>card(id).locator('[data-remove-setup]');
 const buttonStyle=await style(remove(started.id)),neutralStyle=await style(card(neutral.id));
 await stage(started.id,'Started');assert.match(await card(started.id).getAttribute('class'),/gs-domain-started/);assert.notDeepEqual(await style(card(started.id)),neutralStyle);assert.deepEqual(await style(remove(started.id)),buttonStyle);
 assert.equal((await data()).summary.started,baseline.summary.started+1);assert.equal((await data()).summary.todo,baseline.summary.todo+5);
 await page.selectOption('#gs-filter-stage','Started');assert.equal(await page.locator('[data-setup]').count(),1);await page.selectOption('#gs-filter-stage','');
 await stage(complete.id,'Setup complete');assert.match(await card(complete.id).getAttribute('class'),/gs-domain-complete/);assert.deepEqual(await style(remove(complete.id)),buttonStyle);
 await stage(burned.id,'Burned');await stage(qr.id,'Failed at QR');
 for(const job of [burned,qr]){assert.match(await card(job.id).getAttribute('class'),/gs-domain-danger/);assert.deepEqual(await style(remove(job.id)),buttonStyle);}
 const shades=await Promise.all([started,complete,burned].map(j=>style(card(j.id))));assert.equal(new Set(shades.map(s=>s.background)).size,3);assert.deepEqual(await style(card(burned.id)),await style(card(qr.id)));
 model=await data();assert.equal(model.summary.todo,baseline.summary.todo+4);assert.equal(model.summary.burned,baseline.summary.burned+1);assert.equal(model.summary.started,baseline.summary.started+1);assert.equal(model.summary.complete,baseline.summary.complete+1);assert.deepEqual(model.client_progress[started.client_id],{target:5,complete:1,actual_todo:3,todo:4,remaining:4,replacements_needed:1});
 await page.locator('#gs-setup-list').screenshot({path:'work/gmail-stage-shades-desktop.png',animations:'disabled'});
 await page.reload();await idle();assert.match(await card(started.id).getAttribute('class'),/gs-domain-started/);
 await stage(started.id,'Not started');assert.deepEqual(await style(card(started.id)),neutralStyle);await stage(started.id,'Purchase made');assert.match(await card(started.id).getAttribute('class'),/gs-domain-started/);
 await stage(qr.id,'Failed at payment');assert.match(await card(qr.id).getAttribute('class'),/gs-domain-danger/);
 await page.setViewportSize({width:390,height:844});assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false);await page.locator('#gs-stats').screenshot({path:'work/gmail-started-stats-mobile.png',animations:'disabled'});await card(qr.id).screenshot({path:'work/gmail-qr-failed-card-mobile.png',animations:'disabled'});
 assert.deepEqual(errors,[]);await browser.close();console.log('PASS: Started selection/filter/counts, amber progress cards, green completion, red burned/QR/payment failure cards, neutral reset, stable button styles, reload and mobile.');
})().catch(e=>{console.error(e);process.exit(1)});
