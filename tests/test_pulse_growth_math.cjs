const assert=require('node:assert/strict');
const calculate=require('../app/static/pulse-growth-math.js');
function data(changes={}){return {mode:'monthly',measurement_start:'2026-01-01',annual_churn:0,current_revenue:1000,goal_12:2200,goal_24:3400,monthly_leads:100,conversions:[{landed:'2026-02-01',amount:100}],leads:{start:'2026-01-01',as_of:'2026-10-21',count:100},...changes};}
let r=calculate(data());assert.equal(r.at12,2200);assert.equal(r.at24,3400);assert.equal(r.needed12,100);assert.equal(r.needed24,100);assert.equal(r.showRate,false);assert.equal(r.leadsPerSale,100);
// Existing revenue already includes actual conversions: no double count at month zero.
assert.equal(r.timeline[0],1000);
r=calculate(data({mode:'project',current_revenue:1200,goal_12:2400,goal_24:2400}));assert.equal(r.at12,2400);assert.equal(r.at24,2400);assert.equal(r.timeline[24],2400);
// Annual churn is compounded, not 10/12 percentage points added each month.
r=calculate(data({annual_churn:10,monthly_leads:0}));assert(Math.abs(r.at12-900)<1e-8);assert(Math.abs(r.at24-810)<1e-8);
r=calculate(data({annual_churn:100,monthly_leads:100}));assert.equal(r.at12,100);assert.equal(r.at24,100);
// Project sales never recur or compound. Churn does not erase a one-off newly landed sale.
r=calculate(data({mode:'project',current_revenue:0,annual_churn:100}));assert.equal(r.at12,1200);assert.equal(r.at24,1200);
r=calculate(data({monthly_leads:50}));assert.equal(r.sales,12);assert.equal(r.at12,1600);
for(const count of [34,100]) assert.equal(calculate(data({leads:{start:'2026-01-01',as_of:'2026-10-21',count}})).showRate,false);
const three=Array.from({length:3},()=>({landed:'2026-02-01',amount:100}));assert.equal(calculate(data({conversions:three})).showRate,true);
assert.equal(calculate(data({leads:null})).ready,false);assert.equal(calculate(data({leads:{count:0}})).ready,false);assert.equal(calculate(data({conversions:[]})).ready,false);
assert.equal(calculate(data({measurement_start:'2026-02-01'})).ready,false);
assert.equal(calculate(data({conversions:[{landed:'2026-11-01',amount:100}]})).ready,false);
assert.equal(calculate(data({goal_12:0})).needed12,0);
assert.equal(calculate(data({goal_12:1e15})).needed12,null);
// The minimum suggested lead count actually reaches the target under the same model.
for(const mode of ['monthly','project'])for(const churn of [0,25,100]) {
 const d=data({mode,annual_churn:churn,goal_12:5341,goal_24:12345});const out=calculate(d);
 for(const [n,key,goal] of [[out.needed12,'at12',d.goal_12],[out.needed24,'at24',d.goal_24]])if(n!==null){assert(calculate({...d,monthly_leads:n})[key]>=goal);if(n>0)assert(calculate({...d,monthly_leads:n-1})[key]<goal);}
}
// Increasing leads must not reduce a churn-adjusted forecast.
for(const mode of ['monthly','project']) {
 let previous=0;
 for(let leads=0;leads<200;leads++){const out=calculate(data({mode,annual_churn:90,monthly_leads:leads}));assert(out.at24>=previous);previous=out.at24;}
}
console.log('PASS: monthly/project forecasts, annual churn, expected sales, reciprocal metrics, 3% threshold, date alignment and minimum lead targets.');

// Full and partial months use the same observed window as the sales metric.
r=calculate(data({leads:{start:'2026-01-01',as_of:'2026-02-28',count:100}}));assert.equal(r.averageMonthlyLeads,50);
r=calculate(data({leads:{start:'2026-01-01',as_of:'2026-01-15',count:30},conversions:[{landed:'2026-01-02',amount:100}]}));assert(Math.abs(r.averageMonthlyLeads-62)<1e-8);

assert.equal(calculate(data({conversions:[],leads:{start:'2026-01-01',as_of:'2026-01-31',count:0}})).averageMonthlyLeads,0);
