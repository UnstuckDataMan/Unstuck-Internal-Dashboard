// Every provider call uses the synthetic DNS zones in gmail_setup_preview.
const {chromium}=require('playwright'),assert=require('node:assert/strict');
(async()=>{
 const browser=await chromium.launch({executablePath:'/usr/bin/chromium',headless:true,args:['--no-sandbox']});
 const page=await browser.newPage({viewport:{width:1440,height:1050}}),errors=[];
 page.on('pageerror',e=>errors.push(e.message));
 await page.route('**/*',r=>r.request().url().startsWith('http://127.0.0.1:8766')?r.continue():r.abort());
 const idle=()=>page.waitForFunction(()=>!document.querySelector('#gs-app').hasAttribute('aria-busy'));
 const data=()=>page.evaluate(()=>fetch('/api/gmail-setup').then(r=>r.json()));
 const open=async id=>{await page.click(`[data-setup-dns="${id}"]`);await page.waitForFunction(()=>document.querySelector('#gs-dialog').open);await idle();};
 await page.goto('http://127.0.0.1:8766/gmail-set');await idle();
 const owner=(await data()).state.members.find(m=>m.name==='Dylan').id;
 await page.click('#gs-add-setup');await page.fill('#gs-client-name','One click setup journey');await page.selectOption('#gs-client-manager',owner);await page.click('#gs-dialog-save');await page.waitForFunction(()=>!document.querySelector('#gs-dialog').open);await idle();
 await page.fill('#gs-bulk-sender','me');await page.fill('#gs-bulk-domains','cleandnsclient.com\nkeepdnsclient.com');await page.locator('#gs-bulk-panel button[type="submit"]').click();await idle();
 let model=await data();const jobs=model.state.setups.filter(j=>['cleandnsclient.com','keepdnsclient.com'].includes(j.domain)),first=jobs.find(j=>j.domain==='cleandnsclient.com'),neighbor=jobs.find(j=>j.domain==='keepdnsclient.com'),summary=model.summary;
 await open(first.id);assert.match(await page.textContent('#gs-dialog'),/Only proceed if you’ve completed purchase/);assert.equal(await page.textContent('#gs-dialog-save'),'Proceed');assert.equal(await page.locator('.gs-cleanup-preview li').count(),5);
 const text=await page.textContent('.gs-cleanup-preview');assert.match(text,/fwd1.porkbun.com/);assert.match(text,/fwd2.porkbun.com/);assert.match(text,/_spf.porkbun.com/);assert.match(text,/_acme-challenge/);assert.match(await page.textContent('#gs-dialog'),/t.cleandnsclient.com/);assert.match(await page.textContent('#gs-dialog'),/open.sleadtrack.com/);assert.match(await page.textContent('#gs-dialog'),/rua=mailto:me@cleandnsclient.com/);assert.doesNotMatch(text,/smtp.google.com|DMARC/);
 await page.click('#gs-dialog-cancel');assert.equal((await data()).state.setups.find(j=>j.id===first.id).dns_setup,undefined);
 await open(first.id);assert.equal(await page.locator('.gs-cleanup-preview li').count(),5);await page.fill('#gs-dns-destination','clientwebsite.com');await page.click('#gs-dialog-save');await page.waitForFunction(()=>!document.querySelector('#gs-dialog').open);await idle();
 model=await data();assert.deepEqual(model.summary,summary);assert.equal(model.state.setups.find(j=>j.id===first.id).dns_setup.removed,5);assert.equal(model.state.setups.find(j=>j.id===first.id).dns_setup.destination,'https://clientwebsite.com');assert.equal(model.state.setups.find(j=>j.id===neighbor.id).dns_setup,undefined);assert.match(await page.textContent(`[data-setup="${first.id}"]`),/DNS ready/);
 await page.reload();await idle();assert.match(await page.textContent(`[data-setup="${first.id}"]`),/DNS ready/);
 await open(first.id);assert.equal(await page.locator('.gs-cleanup-preview li').count(),0);assert.match(await page.textContent('#gs-dialog'),/No default records to remove/);await page.click('#gs-dialog-cancel');
 await open(neighbor.id);assert.equal(await page.inputValue('#gs-dns-destination'),'https://clientwebsite.com');assert.equal(await page.locator('.gs-cleanup-preview li').count(),5);
 await page.setViewportSize({width:390,height:844});assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false);await page.locator('#gs-dialog').screenshot({path:'work/gmail-dns-cleanup-mobile.png',animations:'disabled'});
 await page.click('#gs-dialog-cancel');assert.deepEqual(errors,[]);await browser.close();console.log('PASS: per-domain DNS cleanup preview, purchase Proceed/Cancel, cancel keeps records, five reviewed deletions, neighboring domain retained, status persistence, repeat readback, unchanged setup stats and mobile dialog.');
})().catch(e=>{console.error(e);process.exit(1)});
