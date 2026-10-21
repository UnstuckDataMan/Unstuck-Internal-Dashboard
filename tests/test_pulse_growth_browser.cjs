const {chromium}=require('playwright');
const fs=require('fs'),path=require('path'),assert=require('node:assert/strict'),{execFileSync}=require('child_process');
const repo=path.resolve(__dirname,'..');
const fragment=execFileSync(path.join(repo,'.venv/bin/python'),['-c',`from jinja2 import Environment,FileSystemLoader,select_autoescape
e=Environment(loader=FileSystemLoader('app/templates'),autoescape=select_autoescape())
print(e.get_template('partials/pulse_growth.html').render(growth_endpoint='/api/planner',growth_portal=False))`],{cwd:repo,encoding:'utf8'});
const start={staff:true,can_activate:true,csrf:'test-csrf',today:'2026-10-21',version:0,enabled:false,include_in_reports:false,data:{mode:'monthly',currency:'GBP',measurement_start:'2026-01-01',annual_churn:0,current_revenue:10000,goal_12:20000,goal_24:30000,monthly_leads:100,conversions:[{id:'one',name:'Example client',landed:'2026-06-01',amount:1000}],leads:{count:100,start:'2026-01-01',as_of:'2026-10-21',source:'Performance Tracker'}}};
(async()=>{
 const browser=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||'/usr/bin/chromium',headless:true,args:['--no-sandbox']});
 const page=await browser.newPage({viewport:{width:1200,height:1000}});let state=structuredClone(start),posts=[],errors=[];
 page.on('pageerror',e=>errors.push(e.message));
 await page.route('https://growth.test/**',async route=>{
   const url=new URL(route.request().url());
   if(url.pathname==='/api/planner'){
    if(route.request().method()==='POST'){
      const payload=route.request().postDataJSON();posts.push(payload);assert.equal(route.request().headers()['x-growth-csrf'],'test-csrf');
      state={...state,version:state.version+1,data:{...payload.data,leads:state.data.leads}};
      if(payload.action==='activate')Object.assign(state,{enabled:true,include_in_reports:true});
    }
    return route.fulfill({json:state});
   }
   if(url.pathname.startsWith('/static/'))return route.fulfill({path:path.join(repo,'app',url.pathname),contentType:url.pathname.endsWith('.css')?'text/css':'application/javascript'});
   return route.fulfill({contentType:'text/html',body:`<!doctype html><html><head><meta charset="utf-8"><style>body{margin:0;padding:28px;background:#10101f;color:#eeeafa}main{max-width:1100px;margin:auto}</style></head><body><main>${fragment}</main></body></html>`});
 });
 await page.goto('https://growth.test/');await page.locator('[data-gp-form]').waitFor({state:'visible'});
 assert((await page.locator('[data-gp-metrics]').innerText()).includes('100'));
 assert(!(await page.locator('[data-gp-metrics]').innerText()).includes('%'));
 assert((await page.locator('[data-gp-forecast]').innerText()).includes('£22,000'));
 await page.locator('[data-gp-activate]').click();assert(await page.locator('dialog').isVisible());
 await page.locator('[data-gp-confirm]').click();assert.equal(posts.length,0);
 await page.locator('[data-gp-approved]').check();await page.locator('[data-gp-converted]').check();await page.locator('[data-gp-confirm]').click();
 await page.locator('dialog').waitFor({state:'hidden'});assert.equal(posts[0].delivery_approved,true);assert.equal(posts[0].has_conversion,true);
 await page.locator('[name=mode]').selectOption('project');assert.equal(await page.locator('[data-gp-current-label]').innerText(),'Current annual revenue');
 assert((await page.locator('[data-gp-forecast]').innerText()).includes('£22,000')); // £10k baseline + 12 x £1k project sales
 await page.locator('[name=monthly_leads]').fill('200');assert((await page.locator('[data-gp-forecast]').innerText()).includes('£34,000'));
 await page.locator('[data-gp-save]').click();await page.waitForFunction(()=>document.querySelector('[data-gp-save-status]').textContent==='All changes saved');
 await page.reload();await page.locator('[data-gp-form]').waitFor({state:'visible'});assert.equal(await page.locator('[name=monthly_leads]').inputValue(),'200');
 await page.locator('[name=mode]').selectOption('monthly');
 await page.locator('[data-months]').fill('4');
 assert((await page.locator('[data-gp-earned]').innerText()).includes('£4,000'));
 assert.equal(await page.locator('[data-gp-earned] svg circle').count(),5);
 await page.locator('[name=currency]').selectOption('USD');
 assert.equal(await page.locator('[data-amount-label]').innerText(),'Monthly value ($)');
 assert((await page.locator('[data-gp-earned]').innerText()).includes('$4,000'));
 await page.locator('[name=mode]').selectOption('project');
 assert((await page.locator('[data-gp-earned]').innerText()).includes('$1,000'));
 assert(await page.locator('[data-gp-earned] svg rect').count()>0);
 await page.locator('[name=currency]').selectOption('GBP');
 // The percentage appears only once the observed rate reaches 3%.
 await page.locator('[data-gp-add]').click();let row=page.locator('.gp-conversion').last();await row.locator('[data-landed]').fill('2026-07');await row.locator('[data-amount]').fill('1000');
 await page.locator('[data-gp-add]').click();row=page.locator('.gp-conversion').last();await row.locator('[data-landed]').fill('2026-08');await row.locator('[data-amount]').fill('1000');
 assert((await page.locator('[data-gp-metrics]').innerText()).includes('3%'));
 await page.locator('[data-gp-sales]').fill('72');assert.equal(await page.locator('[name=monthly_leads]').inputValue(),'200');
 await page.locator('[name=monthly_leads]').fill('300');assert.equal(await page.locator('[data-gp-sales]').inputValue(),'108');
 // Annual target remains editable without lead data and survives saving.
 await page.locator('[data-gp-sales]').fill('72');
 await page.locator('[data-gp-save]').click();await page.waitForFunction(()=>document.querySelector('[data-gp-save-status]').textContent==='All changes saved');
 await page.reload();await page.locator('[data-gp-form]').waitFor({state:'visible'});
 assert.equal(await page.locator('[data-gp-sales]').inputValue(),'72');
 state.data.leads=null;
 await page.reload();await page.locator('[data-gp-form]').waitFor({state:'visible'});
 assert(await page.locator('[data-gp-sales]').isEnabled());
 await page.locator('[data-gp-sales]').fill('120');
 await page.locator('[data-gp-save]').click();await page.waitForFunction(()=>document.querySelector('[data-gp-save-status]').textContent==='All changes saved');
 assert.equal(state.data.annual_sales,120);
 state.data.leads=start.data.leads;
 await page.reload();await page.locator('[data-gp-form]').waitFor({state:'visible'});
 assert.equal(await page.locator('[data-gp-sales]').inputValue(),'120');
 assert.equal(await page.locator('[name=monthly_leads]').inputValue(),'334');
 await page.screenshot({path:'/workspace/scratch/growth-planner-desktop.png',fullPage:true});
 await page.setViewportSize({width:390,height:844});assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false);
 await page.screenshot({path:'/workspace/scratch/growth-planner-mobile.png',fullPage:true});
 assert.deepEqual(errors,[]);await browser.close();console.log('PASS: activation popup, live forecasts, save/reload, 3% display rule, project labels and mobile layout.');
})();
