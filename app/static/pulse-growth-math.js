/* Small client planner, informed by roi-forecast.js: compounded annual churn,
   expected sales, and project revenue counted once. */
(function(root){
  'use strict';
  function calculate(data) {
    const leads = data.leads;
    if (!leads || leads.start !== data.measurement_start || !Number.isFinite(leads.count) || leads.count <= 0)
      return {ready:false, message:'Refresh generated leads to calculate your forecast.'};
    const sales = data.conversions.filter(c=>c.landed >= leads.start && c.landed <= leads.as_of);
    if (!sales.length) return {ready:false, message:'Record a conversion within the lead measurement period to start forecasting.'};
    if (sales.length > leads.count) return {ready:false, message:'There are more conversions than generated leads. Check the tracking dates and refresh leads.'};
    const values = [data.annual_churn,data.current_revenue,data.goal_12,data.goal_24,data.monthly_leads,...sales.map(c=>c.amount)];
    if (!values.every(v=>Number.isFinite(v) && v>=0) || data.annual_churn>100 || !['monthly','project'].includes(data.mode))
      return {ready:false,message:'Check the business figures and conversion amounts.'};
    const rate = sales.length/leads.count;
    const average = sales.reduce((sum,c)=>sum+c.amount,0)/sales.length;
    if (!(average>0)) return {ready:false,message:'Enter the amount of each landed conversion.'};
    const retention = Math.pow(1-data.annual_churn/100,1/12);
    function project(monthlyLeads) {
      let recurring=data.current_revenue, totalSales=0;
      const months=[];
      for(let m=1;m<=24;m++) {
        // Expected averages keep the forecast proportional to the lead target.
        // Rounding arrivals to whole sales can make higher targets look worse
        // under churn by shifting a sale into an earlier month.
        const won=monthlyLeads*rate;totalSales+=won;
        recurring=recurring*retention+won*average;
        const revenue=data.mode==='monthly' ? recurring : data.current_revenue/12*Math.pow(retention,m)+won*average;
        months.push({month:m,revenue,sales:won});
      }
      const at12=data.mode==='monthly'?months[11].revenue:months.slice(0,12).reduce((s,m)=>s+m.revenue,0);
      const at24=data.mode==='monthly'?months[23].revenue:months.slice(12).reduce((s,m)=>s+m.revenue,0);
      // A project goal is annual revenue: show a trailing 12-month total at each point.
      const timeline=[data.current_revenue,...months.map((month,i)=>data.mode==='monthly'?month.revenue:
        months.slice(Math.max(0,i-11),i+1).reduce((s,m)=>s+m.revenue,0)+Math.max(0,11-i)*data.current_revenue/12)];
      return {at12,at24,timeline,sales:totalSales};
    }
    function needed(goal, key) {
      if(project(0)[key]>=goal)return 0;
      if(project(1000000)[key]<goal)return null;
      let lo=0,hi=1000000;
      while(lo<hi){const mid=Math.floor((lo+hi)/2);if(project(mid)[key]>=goal)hi=mid;else lo=mid+1;}
      return lo;
    }
    return {ready:true, count:sales.length, totalCount:data.conversions.length,
      rate, showRate:rate>=0.03, leadsPerSale:1/rate, average,
      ...project(data.monthly_leads), needed12:needed(data.goal_12,'at12'),needed24:needed(data.goal_24,'at24')};
  }
  if(typeof module!=='undefined'&&module.exports)module.exports=calculate;
  else root.pulseGrowthCalculate=calculate;
})(typeof window!=='undefined'?window:globalThis);
