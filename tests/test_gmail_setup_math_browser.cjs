// Fresh synthetic preview: reproduce 22 phone + 2 PC slots against 29 required.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
(async()=>{
 const browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});
 const page=await browser.newPage({viewport:{width:1440,height:1050}}),errors=[];
 page.on('pageerror',e=>errors.push(e.message));
 await page.route('**/*',r=>r.request().url().startsWith('http://127.0.0.1:8766')?r.continue():r.abort());
 const idle=()=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
 const save=async()=>{await page.click('#gs-dialog-save');await page.waitForFunction(()=>!document.querySelector('#gs-dialog').open);await idle();};
 const data=()=>page.evaluate(()=>fetch('/api/gmail-setup').then(r=>r.json()));
 await page.goto('http://127.0.0.1:8766/gmail-set');await idle();
 await page.click('#gs-tab-devices');
 for(const [phones,pcs,condition] of [['10','0','Unused'],['2','0','Used|1'],['0','2','Unused']]){
  await page.click('#gs-add-devices');await page.selectOption('#gs-device-phones',phones);await page.selectOption('#gs-device-pcs',pcs);await page.selectOption('#gs-device-condition',condition);await save();
 }
 const owner=(await data()).state.members.find(m=>m.name==='Dylan').id;
 await page.click('#gs-tab-setups');await page.click('#gs-add-setup');await page.fill('#gs-client-name','Setup maths QA');await page.selectOption('#gs-client-manager',owner);await save();
 await page.fill('#gs-bulk-domains',Array.from({length:19},(_,i)=>`mathcheck${i}.com`).join('\n'));await page.locator('#gs-bulk-panel button[type="submit"]').click();await idle();
 let model=await data();assert.equal(model.summary.available_setups,24);assert.equal(model.summary.target,29);assert.equal(model.summary.shortfall,3);
 assert.match(await page.textContent('#gs-stats'),/Setup capacity \/ required24\/29/);
 assert.match(await page.textContent('.gs-reserve-calculation'),/22 phone \+ 2 laptop \/ PC = 24 setups/);
 assert.match(await page.textContent('.gs-reserve-calculation'),/Gap: 29 − 24 = 5 setups/);
 assert.match(await page.textContent('.gs-reserve-calculation'),/5 ÷ 2, rounded up = 3 new phones \(6 extra setup slots\)/);
 const clientHash=new URL(page.url()).hash;
 await page.goto('http://127.0.0.1:8766/');await page.waitForFunction(()=>document.querySelector('#gs-home-stock')?.textContent==='Order 3 more phones');
 await page.click('#btn-operations');await page.locator('#gs-home-stock').screenshot({path:'work/gmail-operations-short-order.png'});
 await page.goto('http://127.0.0.1:8766/gmail-set'+clientHash);await idle();
 await page.locator('#gs-stats').screenshot({path:'work/gmail-24-of-29-setup-capacity-desktop.png'});
 await page.setViewportSize({width:390,height:844});assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false);await page.locator('#gs-stats').screenshot({path:'work/gmail-24-of-29-setup-capacity-mobile.png'});
 await page.setViewportSize({width:1440,height:1050});await page.locator('.gs-capacity-settings summary').click();
 await page.fill('#gs-pc-capacity','2');await page.locator('#gs-capacity-form button').click();await idle();
 model=await data();assert.equal(model.summary.available_setups,26);assert.equal(model.summary.target,29);assert.equal(model.summary.shortfall,2);
 await page.fill('#gs-phone-capacity','3');await page.locator('#gs-capacity-form button').click();await idle();
 model=await data();assert.equal(model.summary.available_setups,38);assert.equal(model.summary.shortfall,0);assert.match(await page.textContent('.gs-reserve-calculation'),/Surplus: 38 − 29 = 9 setup slots/);
 await page.goto('http://127.0.0.1:8766/');await page.waitForFunction(()=>document.querySelector('#gs-home-stock')?.textContent==='No new devices needed');
 assert.match(await page.textContent('#gs-home-stock'),/^No new devices needed$/);assert.deepEqual(errors,[]);
 await browser.close();console.log('PASS: 22 phone + 2 PC = 24/29 capacity, five slots short rounds to three phones; clear calculation, changed limits, surplus, Operations and mobile.');
})().catch(e=>{console.error(e);process.exit(1)});
