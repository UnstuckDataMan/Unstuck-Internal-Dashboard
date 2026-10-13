const assert = require('node:assert/strict');
const forecast = require('../app/static/roi-forecast.js');
let seed=20261013;
const random=()=>{seed=(Math.imul(seed,1664525)+1013904223)>>>0;return seed/4294967296;};
const close=(a,b)=>assert.ok(Math.abs(a-b)<=2e-9*Math.max(1,Math.abs(b)),`${a} != ${b}`);
let checks=0;
for(let test=0;test<20000;test++) {
 const input={leads:random()*1000,conversion:random()*100,value:random()*1000000,rate:.01+random()*20000,churn:random()*100,mode:test%2?'project':'retainer'};
 const f=forecast(input), noChurn=forecast({...input,churn:0});
 let accumulated=0,adjustedAccumulated=0;
 for(let i=0;i<5;i++) {
  const row=f.years[i];accumulated+=row.revenue;adjustedAccumulated+=row.adjustedRevenue;
  close(row.accumulated,accumulated);close(row.adjustedAccumulated,adjustedAccumulated);
  close(row.ratio,accumulated/(input.rate*12*(i+1)));
  close(row.revenue,noChurn.years[i].revenue);close(row.accumulated,noChurn.years[i].accumulated);
  assert.ok(row.adjustedRevenue<=row.revenue+1e-5);
  assert.ok(row.revenue>=0 && row.adjustedRevenue>=0);
  close(noChurn.years[i].adjustedRevenue,noChurn.years[i].revenue);
  checks+=8;
 }
 close(f.revenue,f.months.reduce((s,m)=>s+m.revenue,0));close(f.revenue,f.years[0].revenue);
 close(f.net,f.revenue-f.investment);
 const scaled=forecast({...input,value:input.value*2});close(scaled.revenue,f.revenue*2);close(scaled.years[4].adjustedAccumulated,f.years[4].adjustedAccumulated*2);
 checks+=5;
}
// Separate cohort-based reference for a further 1,000 scenarios.
for(let test=0;test<1000;test++) {
 const input={leads:random()*30,conversion:random()*50,value:random()*30000,rate:7990,churn:random()*100,mode:test%2?'retainer':'project'};
 const f=forecast(input), s=input.leads*input.conversion/100, r=Math.pow(1-input.churn/100,1/12);
 const starts=Array.from({length:Math.floor(s*60)},(_,k)=>Math.ceil((k+1)/s)-1);
 let accumulated=0;
 for(let year=0;year<5;year++) {
  let revenue=0,adjusted=0;
  for(let month=year*12;month<(year+1)*12;month++) {
   for(const start of starts) {
    if(input.mode==='retainer' && start<=month) {revenue+=input.value;adjusted+=input.value*Math.pow(r,month-start);}
    if(input.mode==='project' && start===month) {revenue+=input.value;adjusted+=input.value*Math.pow(r,month);}
   }
  }
  accumulated+=revenue;close(f.years[year].revenue,revenue);close(f.years[year].adjustedRevenue,adjusted);close(f.years[year].accumulated,accumulated);checks+=3;
 }
}
const base={leads:3,conversion:5,value:25000,rate:7990,churn:30,mode:'retainer'};
assert.deepEqual(forecast(base).years.map(y=>y.revenue),[150000,700000,1225000,1775000,2300000]);
for(const field of ['leads','conversion','value','rate','churn']) for(const invalid of [NaN,Infinity,-Infinity,-1,null,undefined,'5']) { if(field==='churn' && invalid===undefined) {assert.equal(forecast({...base,churn:undefined}).years[0].adjustedRevenue,150000);} else assert.throws(()=>forecast({...base,[field]:invalid}),RangeError); }
for(const invalid of [100.01,1e10]) for(const field of ['conversion','churn']) { if(field==='churn' && invalid===undefined) {assert.equal(forecast({...base,churn:undefined}).years[0].adjustedRevenue,150000);} else assert.throws(()=>forecast({...base,[field]:invalid}),RangeError); }
for(let month=1;month<=60;month++) for(const delta of [-1e-8,0,1e-8]) {
 const s=(1+delta)/month;
 const f=forecast({...base,leads:s,conversion:100,mode:'project'});
 const year=Math.ceil(month/12)-1;
 const booked=f.years.slice(0,year+1).reduce((sum,r)=>sum+r.revenue/base.value,0);
 const expected=Math.floor(s*12*(year+1)+1e-13);
 assert.equal(booked,expected);
}
console.log(JSON.stringify({seed:20261013,randomForecasts:20000,independentCohortScenarios:1000,financialAssertions:checks,boundaryCases:180,invalidInputCases:39,status:'passed'}));
