const assert=require('node:assert/strict'),calculate=require('../app/static/pulse-growth-math.js');
const d={mode:'monthly',current_revenue:10000,average_value:1000,annual_churn:0,measurement_start:'2026-01-01',goals:[20000,30000,null,null,null],conversions:[{landed:'2026-01-15',amount:1000}],leads:{start:'2026-01-01',as_of:'2026-01-31',count:100}};
let r=calculate(d);assert(r.ready);assert.equal(r.horizon,2);assert.equal(r.forecast[12],22000);assert.equal(r.forecast[24],34000);assert.equal(r.forecast[0],10000);assert.equal(r.goals[0].leadsPerMonth,84);assert.equal(r.goals[0].totalDeals,10);assert.equal(r.leadsPerDeal,100);
r=calculate({...d,goals:[null,null,null,null,720000]});assert.equal(r.horizon,5);assert.equal(r.forecast.length,61);assert.equal(r.goals.length,1);
r=calculate({...d,annual_churn:10});assert(Math.abs(r.baseline[12]-9000)<1e-6);assert(Math.abs(r.baseline[24]-8100)<1e-6);
r=calculate({...d,mode:'project'});assert.equal(r.forecast[12],22000);assert.equal(r.forecast[24],22000);
r=calculate({...d,goals:[null,null,null,null,null]});assert.equal(r.goals.length,0);assert.equal(r.horizon,2);
r=calculate({...d,leads:null});assert(!r.ready);assert.equal(r.forecast,null);assert.equal(r.goals[0].leadsPerMonth,null);
r=calculate({...d,average_value:0});assert.equal(r.goals[0].dealsPerMonth,null);
r=calculate({...d,goals:[0,null,null,null,null]});assert.equal(r.goals[0].leadsPerMonth,0);
r=calculate({...d,annual_churn:100});assert.equal(r.forecast[12],1000);
for(const mode of ['monthly','project'])for(const churn of [0,10,100]){
 const out=calculate({...d,mode,annual_churn:churn});
 assert(out.goals.every(g=>g.leadsPerMonth>=0&&Number.isFinite(g.leadsPerMonth)));
}
console.log('PASS: optional five-year goals, churn, recurring/project revenue, lead requirements, missing data and no historical double count.');
