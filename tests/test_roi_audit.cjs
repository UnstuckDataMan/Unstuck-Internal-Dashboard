const assert = require('node:assert/strict');
const forecast = require('../app/static/roi-forecast.js');
const close = (actual, expected) => assert.ok(Math.abs(actual - expected) <= 1e-8 * Math.max(1, Math.abs(expected)), `${actual} != ${expected}`);
// Independent reference: sum revenue from individual sale cohorts, rather than
// updating the production model's active-client balance month by month.
let cases = 0;
for (const leads of [0,1,5,10,23]) for (const conversion of [0,5,30,100]) for (const churn of [0,10,50,100]) for (const mode of ['project','retainer']) {
  const value = 2375.47, rate = 5490;
  const result = forecast({leads,conversion,churn,mode,value,rate});
  const salesRate = leads * conversion / 100;
  const saleMonths = salesRate === 0 ? [] : Array.from({length:Math.floor(salesRate * 60 + 1e-12)}, (_,k) => Math.ceil((k + 1) / salesRate - 1e-12) - 1);
  const survival = Math.pow(1 - churn / 100, 1 / 12);
  const reference = Array.from({length:60},(_,m)=>{
    const paidSales = saleMonths.filter(s=>mode === 'project' ? s === m : s <= m);
    const baseline = paidSales.length * value;
    const adjusted = mode === 'project' ? baseline * Math.pow(survival,m) : paidSales.reduce((sum,s)=>sum + value * Math.pow(survival,m - s),0);
    return {baseline,adjusted};
  });
  let cumulative=0;
  for (let y=0;y<5;y++) {
    const rows=reference.slice(y*12,(y+1)*12);
    const annual=rows.reduce((sum,r)=>sum+r.baseline,0);
    cumulative+=annual;
    close(result.years[y].revenue,annual);
    close(result.years[y].accumulated,cumulative);
    close(result.years[y].adjustedRevenue,rows.reduce((sum,r)=>sum+r.adjusted,0));
    close(result.years[y].ratio,cumulative/(rate*12*(y+1)));
  }
  result.months.forEach((row,index)=>{close(row.revenue,reference[index].baseline);close(row.adjustedRevenue,reference[index].adjusted);assert.ok(row.adjustedRevenue<=row.revenue+1e-7);});
  close(result.revenue,result.years[0].revenue);
  close(result.net,result.revenue-rate*12);
  cases++;
}
const base={leads:1,conversion:100,value:100,rate:10};
assert.equal(forecast({...base,leads:(1-5e-10)/12}).deals,0);
assert.equal(forecast({...base,leads:1/12}).deals,1);
assert.throws(()=>forecast({...base,mode:'unknown'}),RangeError);
assert.throws(()=>forecast({...base,leads:1e20}),RangeError);
console.log(`Independent audit passed: ${cases} scenarios, monthly and five-year totals, churn boundaries, sale threshold and invalid inputs.`);
