const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const elements = {};
const keys = ['horizon','result-period','fees','conservative','optimistic','scenario','scenario-name','pdf','pdf-logo','form','leads','baseline','value','rate','currency','mode','churn-enabled','churn','churn-controls','years','top-period','period','chart','months','timing','currency-code','value-description','result','ratio','deals','revenue','net','summary','error'];
keys.forEach(key => { elements[key] = { value: '', textContent: '', innerHTML: '', checked: false, dataset: {}, style: {}, handlers: {}, setAttribute(key,value) { this[key] = value; }, addEventListener(key,fn) { this.handlers[key] = fn; }, get valueAsNumber() { return Number(this.value); } }; });
Object.entries({horizon:'1',scenario:'baseline',leads:'5.00', baseline:'5',conservative:'3',optimistic:'6', rate:'5490', currency:'GBP', mode:'project', churn:'10'}).forEach(([key,value]) => { elements[key].value = value; });
const symbols = [{},{}];
let printed = 0;
const events = {};
const details = [{open:false}];
elements['pdf-logo'].complete=true;elements['pdf-logo'].naturalWidth=2000;elements['pdf-logo'].naturalHeight=390;
let capturedReport;
const context = {document:{createElement:()=>({getContext:()=>({drawImage:()=>{}}),toDataURL:()=> 'data:image/png;base64,test'}),title:'Expectations Setting',getElementById:id=>elements[id.replace('roi-','')],querySelectorAll:selector=>selector === 'details' ? details : symbols},Intl,Date,Number,String,Math,setTimeout:fn=>fn(),window:{jspdf:{jsPDF:function(){}},createRoiPdf:(data)=>{capturedReport=data;return {save:()=>{printed++;}};},roiForecast:require('../app/static/roi-forecast.js'),print:()=>{printed++;},addEventListener:(name,fn)=>{events[name]=fn;},removeEventListener:name=>{delete events[name];}}};
vm.runInNewContext(fs.readFileSync('app/static/roi-calculator.js','utf8'),context);
const update = elements.form.handlers.input;
assert.equal(elements.ratio.textContent,'—');
elements.value.value='10000';update();
assert.equal(elements.deals.textContent,'3');assert.equal(elements.revenue.textContent,'£30,000.00');
assert.match(elements.chart.innerHTML,/<svg/);
assert.equal((elements.months.innerHTML.match(/<tr>/g)||[]).length,60);
assert.match(elements.chart.innerHTML,/31/);
assert.match(elements.years.innerHTML,/Year 5/);
assert.match(elements.years.innerHTML,/£150,000.00/);
elements.currency.value='USD';elements.currency.handlers.change();
assert.equal(elements.rate.value,'7990');assert.equal(elements.revenue.textContent,'$30,000.00');
elements.rate.value='8000';elements.currency.value='GBP';elements.currency.handlers.change();
assert.equal(elements.rate.value,'5490');
elements.currency.value='USD';elements.currency.handlers.change();assert.equal(elements.rate.value,'8000');
elements.mode.value='retainer';update();assert.equal(elements.revenue.textContent,'$150,000.00');
elements['churn-enabled'].checked=true;update();assert.equal(elements['churn-controls'].hidden,false);assert.equal(elements.revenue.textContent,'$150,000.00');assert.match(elements.chart.innerHTML,/stroke-dasharray/);
elements.churn.value='101';update();assert.equal(elements.ratio.textContent,'—');
elements.mode.value='project';elements.churn.value='10';update();assert.equal(elements.revenue.textContent,'$30,000.00');assert.equal(elements['churn-controls'].hidden,false);
elements.rate.value='0';update();assert.equal(elements.ratio.textContent,'—');
elements.form.handlers.reset();assert.equal(elements.rate.value,'7990');assert.equal(elements.horizon.value,'3');assert.equal(elements['result-period'].textContent,'Over 3 years');
elements.horizon.value='1';elements.horizon.handlers.change();
console.log('UI checks passed: editable currency defaults, currency memory, retainer mode, churn toggle, invalid inputs, chart generation and reset.');

assert.equal(elements.pdf.disabled,false);
elements.pdf.handlers.click();
assert.equal(printed,1);
assert.equal(capturedReport.inputs.currency,'USD');
assert.equal(capturedReport.inputs.rate,7990);
assert.equal(capturedReport.inputs.horizon,1);
assert.equal(details[0].open,false);
assert.equal(context.document.title,'Expectations Setting');
elements.value.value='';update();assert.equal(elements.pdf.disabled,true);
elements.pdf.handlers.click();assert.equal(printed,1);
console.log('PDF checks passed: current assumptions, valid-input requirement and direct download without browser headers.');
// Repeated editing exercises stale-result clearing and export snapshots.
for (let i=0;i<1000;i++) {
  elements.currency.value=i%2?'USD':'GBP';elements.currency.handlers.change();
  elements.leads.value=String((i%37)/3);
  elements.baseline.value=String(i%101);
  elements.value.value=String((i+1)*137.25);
  elements.rate.value=String(100+i);
  elements.mode.value=i%2?'retainer':'project';
  elements['churn-enabled'].checked=i%3!==0;elements.churn.value=String(i%101);
  update();
  const f=context.window.roiForecast({leads:Number(elements.leads.value),conversion:Number(elements.baseline.value),value:Number(elements.value.value),rate:Number(elements.rate.value),mode:elements.mode.value,churn:elements['churn-enabled'].checked?Number(elements.churn.value):0});
  const money=new Intl.NumberFormat(elements.currency.value==='USD'?'en-US':'en-GB',{style:'currency',currency:elements.currency.value});
  assert.equal(elements.revenue.textContent,money.format(f.revenue));
  assert.equal(elements.pdf.disabled,false);
  assert.ok(!/NaN|Infinity|undefined/.test(elements.chart.innerHTML+elements.years.innerHTML));
  if(i%50===0) {elements.pdf.handlers.click();assert.equal(capturedReport.forecast.revenue,f.revenue);assert.equal(capturedReport.inputs.currency,elements.currency.value);}
  elements.baseline.value='101';update();assert.equal(elements.pdf.disabled,true);assert.equal(elements.revenue.textContent,'—');
}
console.log('Repeated control stress passed: 1,000 edits, 1,000 invalid-input transitions and 20 PDF snapshots.');

// Scenario defaults and editable per-scenario memory use the unchanged forecast model.
elements.currency.value='GBP';elements.currency.handlers.change();
elements['churn-enabled'].checked=false;elements.value.value='10000';elements.leads.value='5';elements.rate.value='5490';elements.mode.value='project';elements.baseline.value='5';update();
elements.scenario.value='conservative';elements.scenario.handlers.change();assert.equal(elements.conservative.value,'3');assert.equal(elements.revenue.textContent,'£10,000.00');
elements.conservative.value='2.5';update();
elements.scenario.value='optimistic';elements.scenario.handlers.change();assert.equal(elements.optimistic.value,'6');assert.equal(elements.revenue.textContent,'£30,000.00');
elements.scenario.value='conservative';elements.scenario.handlers.change();assert.equal(elements.conservative.value,'2.5');
elements.scenario.value='baseline';elements.scenario.handlers.change();assert.equal(elements.baseline.value,'5');const beforeExport=printed;elements.pdf.handlers.click();assert.equal(printed,beforeExport+1);assert.equal(capturedReport.inputs.scenario,'Probable');assert.equal(capturedReport.inputs.conversion,5);
console.log('Scenario switching and editable rate memory passed.');
elements.scenario.value='baseline';elements.scenario.handlers.change();const baselineChart=elements.chart.innerHTML;
elements.scenario.value='optimistic';elements.scenario.handlers.change();assert.notEqual(elements.chart.innerHTML,baselineChart);assert.equal(elements['scenario-name'].textContent,'Optimistic · 6%');
elements.optimistic.value='12';update();assert.equal(elements.revenue.textContent,'£70,000.00');
elements.baseline.value='101';update();assert.equal(elements.revenue.textContent,'£70,000.00');
elements.scenario.value='baseline';elements.scenario.handlers.change();assert.equal(elements.pdf.disabled,true);
console.log('Visible inputs passed: selected-rate isolation, heading updates, edits and equal-first-year/different-chart case.');

// Headers and values must move together, including when churn and cumulative
// revenue are very different (the reported £150,000/year, 50% churn example).
elements.currency.value='GBP';elements.currency.handlers.change();
elements.scenario.value='baseline';elements.baseline.value='5';elements.leads.value='5';elements.value.value='50000';elements.rate.value='5490';elements.mode.value='project';elements['churn-enabled'].checked=true;elements.churn.value='50';update();
const rows=Array.from(elements.years.innerHTML.matchAll(/<tr>(.*?)<\/tr>/g),m=>Array.from(m[1].matchAll(/<td>(.*?)<\/td>/g),cell=>cell[1]));
assert.equal(rows[0][2],'£101,902.39');assert.equal(rows[0][3],'£150,000.00');
assert.equal(rows[1][2],'£50,951.20');assert.equal(rows[1][3],'£300,000.00');
assert.equal(rows[4][2],'£6,368.90');assert.equal(rows[4][3],'£750,000.00');
const template=fs.readFileSync('app/templates/roi_calculator.html','utf8');
assert.ok(template.indexOf('<th>After churn')<template.indexOf('<th>Total added revenue'));
assert.match(template,/roi-calculator.js\?v=expectations-2/);
console.log('Reported column mismatch regression passed: churn in column 3, accumulated revenue in column 4, and updated script URLs.');

// The headline period changes cumulative revenue, sales and fees together.
// Retainers must include revenue carried forward from earlier acquired clients.
elements['churn-enabled'].checked=false;elements.value.value='10000';elements.leads.value='5';elements.baseline.value='5';elements.rate.value='5490';
for (const [mode, horizon, revenue, fees, deals] of [
  ['project',3,90000,197640,9], ['project',5,150000,329400,15],
  ['retainer',3,1530000,197640,9], ['retainer',5,4350000,329400,15],
  ['retainer',1,150000,65880,3],
]) {
  elements.mode.value=mode;elements.horizon.value=String(horizon);elements.horizon.handlers.change();
  const money=new Intl.NumberFormat('en-GB',{style:'currency',currency:'GBP'});
  assert.equal(elements.revenue.textContent,money.format(revenue));
  assert.equal(elements.fees.textContent,money.format(fees));
  assert.equal(elements.net.textContent,money.format(revenue-fees));
  assert.equal(elements.deals.textContent,String(deals));
  assert.equal(elements.ratio.textContent,(revenue/fees).toLocaleString('en-GB',{minimumFractionDigits:2,maximumFractionDigits:2})+' : 1');
  assert.equal(elements['result-period'].textContent,`Over ${horizon} ${horizon===1?'year':'years'}`);
  assert.equal((elements.months.innerHTML.match(/<tr>/g)||[]).length,60);
  assert.equal((elements.years.innerHTML.match(/<tr>/g)||[]).length,5);
  elements.pdf.handlers.click();
  assert.equal(capturedReport.inputs.horizon,horizon);
  assert.match(elements['top-period'].textContent,new RegExp(capturedReport.labels[horizon*12-1]));
}
for (const invalidHorizon of ['', '0', '2', '6', '3.0', 'NaN']) {
  elements.horizon.value=invalidHorizon;elements.horizon.handlers.change();
  assert.equal(elements.pdf.disabled,true);assert.equal(elements.horizon['aria-invalid'],'true');
  for (const key of ['ratio','revenue','fees','net','deals']) assert.equal(elements[key].textContent,'—');
  assert.match(elements.error.textContent,/1, 3 or 5 years/);
}
elements.form.handlers.reset();
assert.equal(elements.horizon.value,'3');assert.equal(elements.pdf.disabled,false);assert.equal(elements.horizon['aria-invalid'],'false');
assert.equal(elements.revenue.textContent,'£1,530,000.00');assert.equal(elements.fees.textContent,'£197,640.00');
assert.match(template,/<option value="3" selected>/);
console.log('Horizon checks passed: 1/3/5-year cumulative totals, retained client revenue, fee alignment, exports, invalid values and three-year reset.');
