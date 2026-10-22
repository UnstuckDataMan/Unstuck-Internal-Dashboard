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

 assert.equal(await page.locator('[data-goal="1"]').inputValue(),'240,000');
 assert.equal(await page.locator('[data-goal="2"]').inputValue(),'360,000');
 assert.equal(await page.locator('[data-goal="3"]').inputValue(),'');
 assert.equal(await page.locator('[name=average_value]').inputValue(),'12,000');
 assert((await page.locator('[data-gp-forecast]').innerText()).includes('Year 2'));
 assert(!(await page.locator('[data-gp-forecast]').innerText()).includes('Year 3'));
 assert((await page.locator('[data-gp-metrics]').innerText()).includes('Leads per deal'));
 assert(!(await page.locator('[data-gp-metrics]').innerText()).includes('%'));
 assert((await page.locator('[data-gp-analysis]').innerText()).includes('Year 1'));
 await page.locator('[data-gp-focus]').selectOption('2');assert((await page.locator('[data-gp-analysis]').innerText()).includes('Year 2'));
 await page.locator('[data-goal="5"]').fill('1000000');assert((await page.locator('[data-gp-forecast]').innerText()).includes('Year 5'));
 await page.locator('[data-gp-save]').click();await page.waitForFunction(()=>document.querySelector('[data-gp-save-status]').textContent==='All changes saved');
 assert.equal(state.data.schema_version,4);assert.equal(state.data.goals[2],null);assert.equal(state.data.goals[4],1000000);
 await page.reload();await page.locator('[data-gp-form]').waitFor({state:'visible'});assert.equal(await page.locator('[data-goal="5"]').inputValue(),'1,000,000');
 await page.locator('[data-gp-activate]').click();await page.locator('[data-gp-confirm]').click();assert(await page.locator('dialog').isVisible());
 await page.locator('[data-gp-approved]').check();await page.locator('[data-gp-converted]').check();await page.locator('[data-gp-confirm]').click();await page.locator('dialog').waitFor({state:'hidden'});
 assert(state.enabled);
 assert((await page.locator('body').innerText()).includes('Unstuck activity should work alongside'));
 await page.locator('[name=current_revenue]').fill('50,000.50');
 await page.locator('[name=currency]').selectOption('USD');
 assert.equal(await page.locator('[name=current_revenue]').inputValue(),'50,000.50');
 assert((await page.locator('[data-gp-average-label]').innerText()).includes('($)'));
 await page.locator('[data-gp-save]').click();await page.waitForFunction(()=>document.querySelector('[data-gp-save-status]').textContent==='All changes saved');
 assert.equal(state.data.current_revenue,50000.5);assert.equal(state.data.currency,'USD');
 await page.locator('[name=current_revenue]').fill('1,2');
 assert.equal(await page.locator('[name=current_revenue]').evaluate(el=>el.validity.valid),false);
 await page.locator('[name=current_revenue]').fill('50,000');await page.locator('[name=currency]').selectOption('GBP');
 for(const version of [2,3]){
  state.data={...structuredClone(start.data),schema_version:version,average_value:1000,goals:version===2?[240000,360000,null,null,null]:[20000,30000,null,null,null]};
  await page.reload();await page.locator('[data-gp-form]').waitFor({state:'visible'});
  assert.equal(await page.locator('[name=current_revenue]').inputValue(),'120,000');
  assert.equal(await page.locator('[data-goal="1"]').inputValue(),'240,000');
  assert.equal(await page.locator('[data-amount]').inputValue(),'12,000');
 }
 await page.screenshot({path:'/workspace/scratch/cro-simple-desktop.png',fullPage:true});
 await page.addStyleTag({content:'main{max-width:782px}'});
 const tops=await page.locator('.gp-business input').evaluateAll(els=>els.slice(0,2).map(el=>el.getBoundingClientRect().top));
 assert.equal(tops[0],tops[1]);
 const goalTops=await page.locator('[data-goal]').evaluateAll(els=>els.map(el=>el.getBoundingClientRect().top));
 assert.equal(new Set(goalTops).size,1);
 await page.screenshot({path:'/workspace/scratch/cro-report-layout.png',fullPage:true});
 await page.addStyleTag({content:'main{max-width:560px}'});
 const rows=await page.locator('[data-goal]').evaluateAll(els=>els.map(el=>el.getBoundingClientRect().top));
 assert.equal(rows[0],rows[2]);assert.equal(rows[3],rows[4]);assert(rows[3]>rows[2]);
 await page.setViewportSize({width:390,height:844});assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth),false);
 await page.screenshot({path:'/workspace/scratch/cro-simple-mobile.png',fullPage:true});
 assert.deepEqual(errors,[]);await browser.close();console.log('PASS: legacy conversion, optional goals, five-year chart, goal selection, saving, activation and mobile layout.');
})();
