(function(){
'use strict';
function init(root){
 if(root.dataset.ready)return;root.dataset.ready='true';
 const q=s=>root.querySelector(s),form=q('[data-gp-form]'),dialog=q('[data-gp-dialog]');
 let state=null,busy=false,dirty=false,lastActivity=Date.now();
 const fields=['mode','currency','measurement_start','annual_churn','current_revenue','average_value'];
 const field=n=>form.elements.namedItem(n);
 const num=(n,d=1)=>n.toLocaleString('en-GB',{maximumFractionDigits:d});
 const money=n=>new Intl.NumberFormat('en-GB',{style:'currency',currency:field('currency').value,maximumFractionDigits:0}).format(n);
 const message=(s,error=false)=>{q('[data-gp-message]').textContent=s;q('[data-gp-message]').dataset.error=String(error);};
 const markDirty=()=>{lastActivity=Date.now();dirty=true;q('[data-gp-save-status]').textContent='Unsaved changes';};
 const numeric=value=>Number(String(value).replace(/,/g,''));
 function validateMoney(el){
  const value=el.value.trim();
  const valid=!value||/^(?:[0-9]+|[0-9]{1,3}(?:,[0-9]{3})+)(?:\.[0-9]{1,2})?$/.test(value);
  const n=numeric(value),minimum=el.hasAttribute('data-amount')?0.01:0;
  el.setCustomValidity(valid&&(!value||(n>=minimum&&n<=1000000000))?'':'Enter a valid amount, such as 50,000 or 50,000.50.');
 }
 function formatMoneyInput(el){
  validateMoney(el);
  if(el.value.trim()&&el.validity.valid)el.value=numeric(el.value).toLocaleString('en-GB',{minimumFractionDigits:numeric(el.value)%1?2:0,maximumFractionDigits:2});
 }
 function normalized(d){
  if(d.schema_version===4)return d;
  const recurring=d.mode!=='project',sales=d.conversions||[],round=n=>Math.round(n*100)/100;
  const current=d.schema_version===2||d.schema_version===3?round(d.current_revenue*12):recurring?round(d.current_revenue*12):d.current_revenue;
  const average=d.schema_version===2||d.schema_version===3?round(d.average_value*(recurring?12:1)):sales.length?round(sales.reduce((sum,c)=>sum+c.amount,0)/sales.length*(recurring?12:1)):0;
  const goals=d.schema_version===2?d.goals:d.schema_version===3?d.goals.map(v=>v===null?null:round(v*12)):[d.goal_12?round(d.goal_12*(recurring?12:1)):null,d.goal_24?round(d.goal_24*(recurring?12:1)):null,null,null,null];
  return {...d,schema_version:4,current_revenue:current,average_value:average,goals,
   conversions:sales.map(c=>({...c,amount:round(c.amount*(recurring?12:1))}))};
 }
 function row(c={}){
  const el=document.createElement('div');el.className='gp-conversion gp-simple-deal';el.dataset.id=c.id||crypto.randomUUID();el.dataset.originalLanded=c.landed||'';el.dataset.endMonth=c.end_month||'';
  el.innerHTML='<label>Deal / customer <small>optional</small><input data-name maxlength="120"></label><label>Month won<input data-landed type="month" required></label><label><span data-amount-label>Annual value</span><input data-amount data-money type="text" inputmode="decimal" required></label><button type="button">Remove</button>';
  el.querySelector('[data-name]').value=c.name||'';el.querySelector('[data-landed]').value=(c.landed||'').slice(0,7);el.querySelector('[data-landed]').max=state.today.slice(0,7);el.querySelector('[data-amount]').value=c.amount??'';
  el.querySelector('button').onclick=()=>{el.remove();markDirty();render();};q('[data-gp-conversions]').appendChild(el);
 }
 function data(){
  const d={schema_version:4,goal_12:0,goal_24:0,monthly_leads:0,annual_sales:0,target_basis:'leads'};
  fields.forEach(n=>d[n]=['mode','currency','measurement_start'].includes(n)?field(n).value:numeric(field(n).value));
  d.goals=Array.from(root.querySelectorAll('[data-goal]')).map(el=>el.value===''?null:numeric(el.value));
  d.conversions=Array.from(q('[data-gp-conversions]').children).map(el=>{
   const month=el.querySelector('[data-landed]').value;
   const landed=el.dataset.originalLanded.slice(0,7)===month?el.dataset.originalLanded:month===d.measurement_start.slice(0,7)?d.measurement_start:month+'-01';
   return {id:el.dataset.id,name:el.querySelector('[data-name]').value,landed,amount:numeric(el.querySelector('[data-amount]').value),end_month:el.dataset.endMonth};
  });d.leads=state.data.leads;return d;
 }
 function apply(next){
  next.data=normalized(next.data);state=next;fields.forEach(n=>field(n).value=next.data[n]??0);
  root.querySelectorAll('[data-goal]').forEach((el,i)=>el.value=next.data.goals[i]??'');
  field('measurement_start').max=next.today;field('measurement_start').readOnly=!next.staff;
  q('[data-gp-conversions]').replaceChildren();next.data.conversions.forEach(row);
  q('[data-gp-staff]').hidden=!next.staff;q('[data-gp-portal-note]').hidden=next.staff;
  q('[data-gp-activate]').hidden=!next.can_activate||next.enabled;q('[data-gp-deactivate]').hidden=!next.can_activate||!next.enabled;
  q('[data-gp-inclusion-label]').hidden=!next.enabled;q('[data-gp-inclusion]').checked=next.include_in_reports;
  q('[data-gp-refresh]').hidden=!next.staff;q('[data-gp-state]').textContent=next.enabled?'Active':'Draft — staff only';
  q('[data-gp-activation-note]').textContent=next.enabled?'CRO Insights is active. It appears at the bottom of this client’s reports when included.':'Record a deal, then activate after delivery manager approval.';
  root.querySelectorAll('[data-money]').forEach(formatMoneyInput);
  form.hidden=false;dirty=false;q('[data-gp-save-status]').textContent='All changes saved';render();
 }
 function chart(r,d){
  const maxValue=Math.max(d.current_revenue,...r.baseline,...(r.forecast||[]),...r.goals.map(g=>g.value),1),rough=maxValue/4,power=10**Math.floor(Math.log10(rough)),step=[1,2,5,10].map(n=>n*power).find(n=>n>=rough),top=Math.ceil(maxValue/step)*step;
  const symbol={GBP:'£',USD:'$',EUR:'€'}[d.currency],axis=n=>symbol+(n>=1e6?num(n/1e6)+'m':n>=1000?num(n/1000)+'k':num(n));
  const x=m=>90+m/(r.horizon*12)*570,y=n=>190-n/top*150;
  let svg='<svg class="gp-chart" viewBox="0 0 700 240" role="img" aria-label="Annual revenue: forecast and selected yearly goals">';
  for(let v=0;v<=top;v+=step)svg+=`<path d="M90 ${y(v)}H660" stroke="#393149"/><text x="80" y="${y(v)+4}" text-anchor="end">${axis(v)}</text>`;
  const points=r.forecast||r.baseline;
  svg+=`<path d="${points.map((v,i)=>(i?'L':'M')+x(i)+','+y(v)).join(' ')}" fill="none" stroke="#b69aff" stroke-width="3" ${r.forecast?'':'stroke-dasharray="5 5"'}/>`;
  for(let yr=0;yr<=r.horizon;yr++){svg+=`<text x="${x(yr*12)}" y="216" text-anchor="middle">${yr?'Year '+yr:'Now'}</text><circle cx="${x(yr*12)}" cy="${y(points[yr*12])}" r="4" fill="#b69aff"><title>${yr?'Year '+yr:'Now'}: ${money(points[yr*12])}</title></circle>`;}
  r.goals.forEach(g=>svg+=`<circle cx="${x(g.year*12)}" cy="${y(g.value)}" r="6" fill="#33d9cd"><title>Year ${g.year} goal: ${money(g.value)}</title></circle>`);
  return svg+'</svg><p class="gp-note">Annual revenue · Purple: '+(r.forecast?'forecast at the observed Unstuck deal pace':'existing revenue after churn; add deal data for a growth forecast')+' · Turquoise: your goals</p>';
 }
 function render(){
  if(!state)return;const d=data(),symbol={GBP:'£',USD:'$',EUR:'€'}[d.currency];
  root.querySelectorAll('[data-goal-label]').forEach((el,i)=>el.textContent='Year '+(i+1)+' annual revenue goal');
  q('[data-gp-goal-note]').textContent='Annual revenue you want the business to reach by the end of each year. Leave any unused years blank. All money figures use annual values.';
  q('[data-gp-average-label]').textContent=(d.mode==='monthly'?'Average annual revenue per customer':'Average value per project')+' ('+symbol+')';
  root.querySelectorAll('[data-amount-label]').forEach(el=>el.textContent=(d.mode==='monthly'?'Annual value':'Project value')+' ('+symbol+')');
  q('[data-gp-count]').textContent=d.conversions.length+(d.conversions.length===1?' deal':' deals');
  q('[data-gp-source]').textContent=d.leads&&d.leads.start===d.measurement_start?`${num(d.leads.count,0)} leads · ${d.leads.start} to ${d.leads.as_of}`:'Refresh leads for the same period as the deals recorded.';
  const host=q('[data-gp-analysis]');
  if(!form.checkValidity()){q('[data-gp-forecast]').textContent='Complete your business figures and deal values to update the plan.';q('[data-gp-metrics]').replaceChildren();host.textContent='Your growth plan will update as you complete the figures above.';return;}
  const r=window.pulseGrowthCalculate(d);
  q('[data-gp-metrics]').innerHTML=`<div class="gp-metric"><strong>${d.conversions.length}</strong>Deals recorded</div>`+(r.ready?`<div class="gp-metric"><strong>${num(r.leadsPerDeal)}</strong>Leads per deal</div><div class="gp-metric"><strong>${money(r.dealValue)}</strong>Average recorded ${d.mode==='monthly'?'annual':'project'} deal value</div>`:'');
  q('[data-gp-forecast]').innerHTML=chart(r,d);
  const select=q('[data-gp-focus]'),previous=select.value;
  select.replaceChildren(...r.goals.map(g=>new Option('Year '+g.year+' — '+money(g.value),String(g.year))));
  if(r.goals.some(g=>String(g.year)===previous))select.value=previous;
  select.disabled=!r.goals.length;select.parentElement.hidden=!r.goals.length;
  if(!r.goals.length){host.textContent='Add a goal for any year to explore the customers and leads needed. You can start with Years 1 and 2 and add later years when ready.';return;}
  const g=r.goals.find(g=>String(g.year)===select.value);
  let copy=`Your Year ${g.year} goal is ${money(g.value)} per year by the end of Year ${g.year}. `;
  if(!g.gap)copy+='Existing revenue is projected to cover this goal under your churn assumption. New customers provide room to grow further.';
  else if(g.dealsPerMonth===null)copy+='Enter an average customer value above zero to calculate the new business needed.';
  else{
   copy+=`Closing the remaining gap would take about ${num(g.dealsPerMonth*12)} new ${d.mode==='monthly'?'customers':'projects'} per year, or ${num(g.totalDeals,0)} over ${g.year} ${g.year===1?'year':'years'}, at your entered average value. `;
   if(r.ready)copy+=`If Unstuck alone filled that gap, the observed leads per deal suggest around ${num(g.leadsPerMonth*12,0)} leads per year (${num(g.leadsPerMonth*g.year*12,0)} in total). `;
   else copy+='Refresh leads and record a deal to translate this into lead requirements. ';
   if(r.ready)copy+=r.forecast[g.year*12]>=g.value?'The observed Unstuck pace supports this goal under these assumptions.':'This gives us a clear planning target to build towards together. Strong follow-up, customer retention and other growth channels can help close the gap.';
  }
  if(r.count!==r.totalCount)copy+=' Only deals within the refreshed lead period are used for leads-per-deal calculations.';
  host.textContent=copy;
 }
    async function load(){
      message('Loading planner…');q('[data-gp-reload]').hidden=true;
      try{const response=await fetch(root.dataset.endpoint,{credentials:'same-origin',cache:'no-store'});const next=await response.json();if(!response.ok){if(response.status===404&&root.dataset.portal==='true'){root.remove();return;}throw Error(next.error||next.detail||'Could not load the planner.');}apply(next);message('');}
      catch(e){message(e.message,true);q('[data-gp-reload]').hidden=false;}
    }
    async function renewSession(){
      // Refresh authentication/CSRF only; never replace unsaved fields or versions.
      const response=await fetch(root.dataset.endpoint,{credentials:'same-origin',cache:'no-store',headers:{Accept:'application/json'}});
      if(response.status===401||response.redirected)throw Error('Please sign in again in another tab, then save here. Your entries are still on this page.');
      const next=await response.json();
      if(!response.ok)throw Error(next.error||next.detail||'Could not reconnect. Your entries are still here; try saving again.');
      state.csrf=next.csrf;
    }
    setInterval(()=>{
      if(state&&!busy&&document.visibilityState==='visible'&&Date.now()-lastActivity<10*60*1000){
        renewSession().catch(e=>message(e.message,true));
      }
    },5*60*1000);
    async function save(action){
      if(busy)return;if(action!=='deactivate'&&!form.reportValidity())return;
      if(action==='activate'&&(!q('[data-gp-approved]').checked||!q('[data-gp-converted]').checked)){q('[data-gp-dialog-error]').textContent='Confirm both questions before activating.';return;}
      if(action==='deactivate'&&!confirm('Deactivate this planner and remove it from client reports?'))return;
      busy=true;root.querySelectorAll('button').forEach(b=>b.disabled=true);message('Saving…');
      const payload={action,version:state.version,data:action==='deactivate'?state.data:data(),include_in_reports:q('[data-gp-inclusion]').checked,delivery_approved:q('[data-gp-approved]').checked,has_conversion:q('[data-gp-converted]').checked};
      // Avoid accepting edits while the in-flight save owns this version.
      form.querySelectorAll('input,select').forEach(el=>el.disabled=true);
      try{await renewSession();const response=await fetch(root.dataset.endpoint,{method:'POST',credentials:'same-origin',headers:{'Content-Type':'application/json','X-Growth-CSRF':state.csrf},body:JSON.stringify(payload)});const next=await response.json();if(!response.ok){if(response.status===409)q('[data-gp-reload]').hidden=false;throw Error(next.error||next.detail||'Could not save.');}apply(next);dialog.close();message(action==='activate'?'Activated. CRO Insights now appears at the bottom of this client’s published reports.':'Saved.');}
      catch(e){message(e.message,true);q('[data-gp-dialog-error]').textContent=e.message;}
      finally{busy=false;root.querySelectorAll('button').forEach(b=>b.disabled=false);form.querySelectorAll('input,select').forEach(el=>el.disabled=false);render();}
    }

 q('[data-gp-focus]').onchange=()=>render();
 form.addEventListener('input',event=>{if(event.target.matches('[data-money]'))validateMoney(event.target);markDirty();render();});
 form.addEventListener('focusout',event=>{if(event.target.matches('[data-money]')){formatMoneyInput(event.target);render();}});
 form.addEventListener('submit',event=>{event.preventDefault();save('save');});
 q('[data-gp-add]').onclick=()=>{row();markDirty();render();};
 q('[data-gp-refresh]').onclick=()=>save('refresh');
 q('[data-gp-activate]').onclick=()=>{if(!form.reportValidity())return;q('[data-gp-approved]').checked=false;q('[data-gp-converted]').checked=false;q('[data-gp-dialog-error]').textContent='';dialog.showModal();};
 q('[data-gp-confirm]').onclick=()=>save('activate');q('[data-gp-cancel]').onclick=()=>dialog.close();
 q('[data-gp-deactivate]').onclick=()=>save('deactivate');
 q('[data-gp-reload]').onclick=()=>{if(!dirty||confirm('Discard unsaved changes and load the saved plan?'))load();};
 window.addEventListener('beforeunload',event=>{if(dirty){event.preventDefault();event.returnValue='';}});
 load();
}
document.querySelectorAll('[data-growth-planner]').forEach(init);
})();
