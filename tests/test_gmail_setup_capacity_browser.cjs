// Synthetic preview only. Run after the existing journey/inventory browsers.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
(async()=>{
 const browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});
 const page=await browser.newPage({viewport:{width:1440,height:1050}}),errors=[];
 page.on('pageerror',e=>errors.push(e.message));page.on('dialog',d=>d.accept());
 await page.route('**/*',r=>r.request().url().startsWith('http://127.0.0.1:8766')?r.continue():r.abort());
 const idle=()=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
 const data=()=>page.evaluate(()=>fetch('/api/gmail-setup').then(r=>r.json()));
 const save=async()=>{await page.click('#gs-dialog-save');await page.waitForFunction(()=>!document.querySelector('#gs-dialog').open);await idle();};
 await page.goto('http://127.0.0.1:8766/gmail-set');await idle();
 let model=await data();const owner=model.state.members.find(m=>m.name==='Robyn').id;
 const phone=model.state.devices.find(d=>d.member_id===owner&&d.kind==='Phone'&&d.use_count===2);
 const pc=model.state.devices.find(d=>d.member_id===owner&&d.kind==='PC'&&d.use_count===1);
 assert.ok(phone&&pc);const card=d=>page.locator(`[data-device="${d.id}"]`);
 await page.click('#gs-tab-devices');
 assert.match(await card(phone).getAttribute('class'),/exhausted phone/);assert.match(await card(pc).getAttribute('class'),/exhausted pc/);
 await page.locator('.gs-capacity-settings summary').click();
 await page.fill('#gs-phone-capacity','3');await page.fill('#gs-pc-capacity','2');await page.locator('#gs-capacity-form button').click();await idle();
 model=await data();assert.equal(model.summary.phone_setups,3);assert.equal(model.summary.pc_setups,2);
 assert.match(await card(phone).getAttribute('class'),/used phone/);assert.match(await card(phone).textContent(),/2 \/ 3 setups/);
 assert.match(await card(pc).getAttribute('class'),/used pc/);assert.match(await card(pc).textContent(),/1 \/ 2 setups/);
 for(const d of [phone,pc])assert.match(await card(d).evaluate(e=>getComputedStyle(e).backgroundColor),/251, 146, 60/);
 for(const d of model.state.devices.filter(d=>d.status==='Blocked'))assert.match(await card(d).getAttribute('class'),/blocked/);
 await page.reload();await idle();assert.equal(await page.inputValue('#gs-phone-capacity'),'3');assert.equal(await page.inputValue('#gs-pc-capacity'),'2');
 assert.equal(await page.locator('#gs-reserve-target').evaluate(e=>e.tagName),'STRONG');
 assert.equal(await page.textContent('#gs-reserve-target'),String((await data()).summary.target));
 await page.click('#gs-add-setup');await page.fill('#gs-client-name','Capacity edits QA');await page.selectOption('#gs-client-manager',owner);await page.selectOption('#gs-client-operator',owner);await save();
 await page.fill('#gs-bulk-sender','me');await page.fill('#gs-bulk-domains','capacityphone.com\ncapacitypc.com');await page.locator('#gs-bulk-panel button[type="submit"]').click();await idle();
 const job=domain=>page.locator('[data-setup]').filter({has:page.locator('h4',{hasText:domain})});
 for(const [domain,d] of [['capacityphone.com',phone],['capacitypc.com',pc]]){
  await job(domain).locator('[data-set-device]').selectOption(d.id);await idle();
  await job(domain).locator('[data-set-stage]').selectOption('Purchase made');await idle();
  await job(domain).locator('[data-set-stage]').selectOption('Setup complete');await idle();
 }
 model=await data();assert.equal(model.state.devices.find(d=>d.id===phone.id).use_count,3);assert.equal(model.state.devices.find(d=>d.id===pc.id).use_count,2);
 await page.click('#gs-tab-devices');assert.match(await card(phone).getAttribute('class'),/exhausted phone/);assert.match(await card(pc).getAttribute('class'),/exhausted pc/);
 for(const d of [phone,pc])assert.match(await card(d).evaluate(e=>getComputedStyle(e).backgroundColor),/248, 113, 113/);
 // Raising limits updates colour and availability again without altering uses.
 await page.locator('.gs-capacity-settings summary').click();await page.fill('#gs-phone-capacity','4');await page.fill('#gs-pc-capacity','3');await page.selectOption('#gs-order-kind','PC');await page.locator('#gs-capacity-form button').click();await idle();
 assert.match(await card(phone).getAttribute('class'),/used phone/);assert.match(await card(pc).getAttribute('class'),/used pc/);
 await page.click('#gs-add-devices');assert.equal(await page.locator('#gs-device-condition option').filter({hasText:'Used 4 times'}).count(),1);await page.click('#gs-dialog-cancel');
 model=await data();assert.equal(model.summary.order_kind,'PC');assert.equal(model.summary.order_capacity,3);
 assert.match(await page.textContent('.gs-reserve-calculation'),/laptops \/ PCs|cover 3 setups/);
 await page.locator('#gs-stats').screenshot({path:'work/gmail-capacity-auto-target-desktop.png'});
 await page.setViewportSize({width:390,height:844});assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false);
 await page.locator('.gs-capacity-settings').screenshot({path:'work/gmail-capacity-settings-mobile.png'});
 await page.locator('#gs-stats').screenshot({path:'work/gmail-capacity-auto-target-mobile.png'});
 await page.goto('http://127.0.0.1:8766/');await page.waitForFunction(()=>document.querySelector('#gs-home-stock')?.textContent.length>0);
 assert.equal(await page.textContent('#gs-home-stock'),'No new devices needed');
 assert.deepEqual(errors,[]);await browser.close();console.log('PASS: shared capacity settings persist, phone 3/PC 2 amber-red thresholds, changing limits updates colours/history/pickers/inventory and Operations, automatic target and mobile layout.');
})().catch(e=>{console.error(e);process.exit(1)});
