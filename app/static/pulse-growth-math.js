(function(root){
'use strict';
function calculate(d){
 const goals=(d.goals||[]).map((value,i)=>({year:i+1,value})).filter(g=>g.value!==null&&g.value!==''&&Number.isFinite(g.value));
 const horizon=goals.length?Math.max(...goals.map(g=>g.year)):2;
 const leads=d.leads,conversions=d.conversions||[];
 const matched=leads&&leads.start===d.measurement_start&&Number.isFinite(leads.count)&&leads.count>=0;
 const sales=matched?conversions.filter(c=>c.landed>=leads.start&&c.landed<=leads.as_of):[];
 let months=0;
 if(matched){
  const from=new Date(leads.start+'T00:00:00Z'),until=new Date(leads.as_of+'T00:00:00Z');
  for(let p=new Date(Date.UTC(from.getUTCFullYear(),from.getUTCMonth(),1));p<=until;){
   const next=new Date(Date.UTC(p.getUTCFullYear(),p.getUTCMonth()+1,1));
   months+=Math.max(0,Math.min(+next,+until+86400000)-Math.max(+p,+from))/(next-p);p=next;
  }
 }
 const ready=matched&&sales.length>0&&leads.count>=sales.length&&months>0;
 const rate=ready?sales.length/leads.count:0,pace=ready?sales.length/months:0;
 const retention=Math.pow(1-d.annual_churn/100,1/12);
 function project(salesPerMonth){
  let revenue=d.current_revenue;const points=[revenue*12];
  for(let m=1;m<=horizon*12;m++){
   revenue=d.mode==='monthly'?revenue*retention+salesPerMonth*d.average_value:
     d.current_revenue*Math.pow(retention,m)+salesPerMonth*d.average_value;
   points.push(revenue*12);
  }return points;
 }
 const baseline=project(0),unit=project(1),forecast=ready?project(pace):null;
 const targets=goals.map(g=>{
  const gap=Math.max(0,g.value-baseline[g.year*12]),perSale=unit[g.year*12]-baseline[g.year*12];
  const dealsPerMonth=gap===0?0:perSale>0?gap/perSale:null;
  const leadsPerMonth=ready&&dealsPerMonth!==null?Math.max(0,Math.ceil(dealsPerMonth/rate-1e-9)):null;
  return {...g,gap,dealsPerMonth,leadsPerMonth,totalDeals:dealsPerMonth===null?null:Math.max(0,Math.ceil(dealsPerMonth*g.year*12-1e-9))};
 });
 return {ready,horizon,goals:targets,forecast,baseline,count:sales.length,totalCount:conversions.length,
  leadsPerDeal:ready?1/rate:null,dealValue:sales.length?sales.reduce((s,c)=>s+c.amount,0)/sales.length:null,
  message:ready?'':'Record at least one deal and refresh its matching lead period to see the forecast.'};
}
if(typeof module!=='undefined'&&module.exports)module.exports=calculate;else root.pulseGrowthCalculate=calculate;
})(typeof window!=='undefined'?window:globalThis);
