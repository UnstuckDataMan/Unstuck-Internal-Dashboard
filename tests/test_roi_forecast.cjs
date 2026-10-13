const assert = require('node:assert/strict');
const forecast = require('../app/static/roi-forecast.js');
const base = { leads: 5, conversion: 5, value: 10000, rate: 5490 };
const project = forecast(base);
assert.deepEqual(project.months.map(row => row.sales), [0,0,0,1,0,0,0,1,0,0,0,1]);
assert.equal(project.deals, 3);
assert.equal(project.revenue, 30000);
assert.equal(project.net, 30000 - 5490 * 12);
const retainer = forecast({ ...base, mode: 'retainer' });
assert.deepEqual(retainer.months.map(row => row.revenue), [0,0,0,10000,10000,10000,10000,20000,20000,20000,20000,30000]);
assert.equal(retainer.revenue, 150000);
const churned = forecast({ ...base, mode: 'retainer', churn: 50 });
assert.ok(churned.years[0].adjustedRevenue < retainer.revenue);
assert.equal(churned.months[3].revenue, 10000);
assert.equal(forecast({ ...base, churn: 50 }).revenue, project.revenue);
assert.equal(forecast({ ...base, mode: 'retainer', churn: 100 }).years[0].adjustedRevenue, project.revenue);
assert.equal(forecast({ ...base, leads: 0 }).revenue, 0);
assert.equal(forecast({ ...base, leads: 20, conversion: 10 }).deals, 24);
assert.equal(forecast({ ...base, leads: 1, conversion: 10 }).deals, 1);
assert.throws(() => forecast({ ...base, rate: 0 }), RangeError);
assert.throws(() => forecast({ ...base, churn: 101 }), RangeError);
assert.throws(() => forecast({ ...base, value: 1e308 }), RangeError);
console.log('Forecast checks passed: spaced sales, project fees, accumulating retainers, churn, zero leads, fractional sales and invalid figures.');

assert.equal(project.years.length, 5);
assert.deepEqual(project.years.map(row => row.revenue), [30000,30000,30000,30000,30000]);
assert.equal(project.years[4].accumulated,150000);
assert.equal(project.years[4].ratio,150000 / (5490 * 60));
assert.equal(retainer.years[1].revenue,510000);
assert.equal(retainer.years[1].accumulated,660000);
assert.equal(retainer.years[0].revenue,retainer.revenue);
assert.ok(churned.years[4].adjustedAccumulated < retainer.years[4].accumulated);
assert.equal(forecast({...base, mode:'retainer', churn:100}).years[4].adjustedAccumulated,150000);
assert.equal(forecast({...base, leads:0}).years[4].accumulated,0);
console.log('Five-year checks passed: cumulative revenue and fees, retainer carryover, churn and first-year consistency.');

assert.deepEqual(churned.years.map(row => row.revenue), retainer.years.map(row => row.revenue));
assert.deepEqual(churned.years.map(row => row.accumulated), retainer.years.map(row => row.accumulated));
assert.deepEqual(churned.years.map(row => row.ratio), retainer.years.map(row => row.ratio));
const projectChurn = forecast({...base, churn:50});
assert.equal(projectChurn.revenue,project.revenue);
assert.ok(projectChurn.years[0].adjustedRevenue < project.revenue);
assert.ok(projectChurn.years[4].adjustedRevenue < projectChurn.years[0].adjustedRevenue);
assert.equal(project.years[4].adjustedRevenue,project.years[4].revenue);
console.log('Separate churn scenarios passed for projects and retainers; baseline revenue and ROI unchanged.');

assert.equal(retainer.timeline.length,60);
for(let year=0;year<5;year++) {
 assert.equal(retainer.timeline.slice(year*12,(year+1)*12).reduce((sum,row)=>sum+row.revenue,0),retainer.years[year].revenue);
}
console.log('60-month chart reconciles with every annual total.');
