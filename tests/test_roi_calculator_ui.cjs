const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const elements = {};
const keys = ['scenario','scenario-name','pdf','pdf-logo','form','leads','conversion','value','rate','currency','mode','churn-enabled','churn','churn-controls','years','top-period','period','chart','months','timing','currency-code','value-description','result','ratio','deals','revenue','net','summary','error'];
keys.forEach(key => { elements[key] = { value: '', textContent: '', innerHTML: '', checked: false, dataset: {}, style: {}, handlers: {}, setAttribute(key,value) { this[key] = value; }, addEventListener(key,fn) { this.handlers[key] = fn; }, get valueAsNumber() { return Number(this.value); } }; });
Object.entries({scenario:'baseline',leads:'5.00', conversion:'5', rate:'5490', currency:'GBP', mode:'project', churn:'10'}).forEach(([key,value]) => { elements[key].value = value; });
const symbols = [{},{}];
let printed = 0;
const events = {};
const details = [{open:false}];
elements['pdf-logo'].complete=true;elements['pdf-logo'].naturalWidth=2000;elements['pdf-logo'].naturalHeight=390;
let capturedReport;
const context = {document:{createElement:()=>({getContext:()=>({drawImage:()=>{}}),toDataURL:()=> 'data:image/png;base64,test'}),title:'ROI Calculator',getElementById:id=>elements[id.replace('roi-','')],querySelectorAll:selector=>selector === 'details' ? details : symbols},Intl,Date,Number,String,Math,setTimeout:fn=>fn(),window:{jspdf:{jsPDF:function(){}},createRoiPdf:(data)=>{capturedReport=data;return {save:()=>{printed++;}};},roiForecast:require('../app/static/roi-forecast.js'),print:()=>{printed++;},addEventListener:(name,fn)=>{events[name]=fn;},removeEventListener:name=>{delete events[name];}}};
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
elements.form.handlers.reset();assert.equal(elements.rate.value,'7990');
console.log('UI checks passed: editable currency defaults, currency memory, retainer mode, churn toggle, invalid inputs, chart generation and reset.');

assert.equal(elements.pdf.disabled,false);
elements.pdf.handlers.click();
assert.equal(printed,1);
assert.equal(capturedReport.inputs.currency,'USD');
assert.equal(capturedReport.inputs.rate,7990);
assert.equal(details[0].open,false);
assert.equal(context.document.title,'ROI Calculator');
elements.value.value='';update();assert.equal(elements.pdf.disabled,true);
elements.pdf.handlers.click();assert.equal(printed,1);
console.log('PDF checks passed: current assumptions, valid-input requirement and direct download without browser headers.');
// Repeated editing exercises stale-result clearing and export snapshots.
for (let i=0;i<1000;i++) {
  elements.currency.value=i%2?'USD':'GBP';elements.currency.handlers.change();
  elements.leads.value=String((i%37)/3);
  elements.conversion.value=String(i%101);
  elements.value.value=String((i+1)*137.25);
  elements.rate.value=String(100+i);
  elements.mode.value=i%2?'retainer':'project';
  elements['churn-enabled'].checked=i%3!==0;elements.churn.value=String(i%101);
  update();
  const f=context.window.roiForecast({leads:Number(elements.leads.value),conversion:Number(elements.conversion.value),value:Number(elements.value.value),rate:Number(elements.rate.value),mode:elements.mode.value,churn:elements['churn-enabled'].checked?Number(elements.churn.value):0});
  const money=new Intl.NumberFormat(elements.currency.value==='USD'?'en-US':'en-GB',{style:'currency',currency:elements.currency.value});
  assert.equal(elements.revenue.textContent,money.format(f.revenue));
  assert.equal(elements.pdf.disabled,false);
  assert.ok(!/NaN|Infinity|undefined/.test(elements.chart.innerHTML+elements.years.innerHTML));
  if(i%50===0) {elements.pdf.handlers.click();assert.equal(capturedReport.forecast.revenue,f.revenue);assert.equal(capturedReport.inputs.currency,elements.currency.value);}
  elements.conversion.value='101';update();assert.equal(elements.pdf.disabled,true);assert.equal(elements.revenue.textContent,'—');
}
console.log('Repeated control stress passed: 1,000 edits, 1,000 invalid-input transitions and 20 PDF snapshots.');

// Scenario defaults and editable per-scenario memory use the unchanged forecast model.
elements.currency.value='GBP';elements.currency.handlers.change();
elements['churn-enabled'].checked=false;elements.value.value='10000';elements.leads.value='5';elements.rate.value='5490';elements.mode.value='project';elements.conversion.value='3';update();
elements.scenario.value='conservative';elements.scenario.handlers.change();assert.equal(elements.conversion.value,'2');assert.equal(elements.revenue.textContent,'£10,000.00');
elements.conversion.value='2.5';update();
elements.scenario.value='optimistic';elements.scenario.handlers.change();assert.equal(elements.conversion.value,'6');assert.equal(elements.revenue.textContent,'£30,000.00');
elements.scenario.value='conservative';elements.scenario.handlers.change();assert.equal(elements.conversion.value,'2.5');
elements.scenario.value='baseline';elements.scenario.handlers.change();assert.equal(elements.conversion.value,'3');const beforeExport=printed;elements.pdf.handlers.click();assert.equal(printed,beforeExport+1);assert.equal(capturedReport.inputs.scenario,'Baseline');assert.equal(capturedReport.inputs.conversion,3);
console.log('Scenario switching and editable rate memory passed.');
