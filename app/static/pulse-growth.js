(function(){
  'use strict';
  function init(root){
    if(root.dataset.ready)return;root.dataset.ready='true';
    const q=s=>root.querySelector(s),form=q('[data-gp-form]'),dialog=q('[data-gp-dialog]');
    let state=null,busy=false,dirty=false;
    const fields=['mode','currency','measurement_start','annual_churn','current_revenue','goal_12','goal_24','monthly_leads'];
    const field=n=>form.elements.namedItem(n);
    function message(text,error=false){q('[data-gp-message]').textContent=text;q('[data-gp-message]').dataset.error=String(error);}
    function money(n){return new Intl.NumberFormat('en-GB',{style:'currency',currency:field('currency').value,maximumFractionDigits:0}).format(n);}
    function num(n,d=1){return n.toLocaleString('en-GB',{maximumFractionDigits:d});}
    function markDirty(){dirty=true;q('[data-gp-save-status]').textContent='Unsaved changes';}
    function row(conversion={}){
      const el=document.createElement('div');el.className='gp-conversion';el.dataset.id=conversion.id||crypto.randomUUID();
      el.innerHTML='<label>Sale / client <small>optional</small><input data-name maxlength="120"></label><label>Date landed<input data-landed type="date" required></label><label><span data-amount-label>Amount</span><input data-amount type="number" min="0.01" max="1000000000" step="0.01" required></label><button type="button" aria-label="Remove conversion">Remove</button>';
      el.querySelector('[data-name]').value=conversion.name||'';
      el.querySelector('[data-landed]').value=conversion.landed||'';
      el.querySelector('[data-landed]').max=state.today;
      el.querySelector('[data-amount]').value=conversion.amount??'';
      el.querySelector('button').onclick=()=>{el.remove();markDirty();render();};
      q('[data-gp-conversions]').appendChild(el);
    }
    function data(){
      const d={};fields.forEach(n=>d[n]=['mode','currency','measurement_start'].includes(n)?field(n).value:Number(field(n).value));
      d.conversions=Array.from(q('[data-gp-conversions]').children).map(el=>({id:el.dataset.id,name:el.querySelector('[data-name]').value,landed:el.querySelector('[data-landed]').value,amount:Number(el.querySelector('[data-amount]').value)}));
      d.leads=state.data.leads;return d;
    }
    function apply(next){
      state=next;fields.forEach(n=>field(n).value=next.data[n]);field('measurement_start').max=next.today;
      q('[data-gp-conversions]').replaceChildren();next.data.conversions.forEach(row);
      q('[data-gp-staff]').hidden=!next.staff;q('[data-gp-portal-note]').hidden=next.staff;
      q('[data-gp-activate]').hidden=!next.can_activate||next.enabled;
      q('[data-gp-deactivate]').hidden=!next.can_activate||!next.enabled;
      q('[data-gp-inclusion-label]').hidden=!next.enabled;q('[data-gp-inclusion]').checked=next.include_in_reports;
      q('[data-gp-refresh]').hidden=!next.staff;field('measurement_start').readOnly=!next.staff;
      q('[data-gp-activation-note]').textContent=next.enabled?'Activated after delivery manager approval. Choose whether this planner is shown in client reports.':'Save your business figures and at least one landed conversion, then activate after delivery manager approval.';
      q('[data-gp-state]').textContent=next.enabled?'Active':'Draft — staff only';q('[data-gp-state]').className='gp-state';
      form.hidden=false;dirty=false;q('[data-gp-save-status]').textContent='All changes saved';render();
    }
    function chart(values,d){
      const max=Math.max(...values,d.goal_12,d.goal_24,1)*1.1;
      const x=i=>48+i/24*624,y=v=>170-v/max*140;
      const path=values.map((v,i)=>(i?'L':'M')+x(i).toFixed(1)+','+y(v).toFixed(1)).join(' ');
      return `<svg class="gp-chart" viewBox="0 0 720 215" role="img" aria-label="Projected revenue over 24 months. Exact 12 and 24 month results are shown above."><path d="M48 30V170H672" fill="none" stroke="#514963"/><path d="${path}" fill="none" stroke="#b69aff" stroke-width="3"/><circle cx="${x(12)}" cy="${y(d.goal_12)}" r="5" fill="#33d9cd"/><circle cx="${x(24)}" cy="${y(d.goal_24)}" r="5" fill="#33d9cd"/><text x="48" y="193">Now</text><text x="333" y="193">12 months</text><text x="608" y="193">24 months</text></svg><p class="gp-note">Purple: forecast · Turquoise: your goals</p>`;
    }
    function render(){
      if(!state)return;const d=data(),monthly=d.mode==='monthly';
      q('[data-gp-current-label]').textContent=monthly?'Current monthly revenue':'Current annual revenue';
      q('[data-gp-goal12-label]').textContent=monthly?'Monthly revenue goal in 12 months':'Revenue goal for the next 12 months';
      q('[data-gp-goal24-label]').textContent=monthly?'Monthly revenue goal in 24 months':'Annual revenue goal for year two';
      q('[data-gp-basis-note]').textContent=monthly?'Enter monthly recurring amounts. Annual churn is spread across the year and applies to existing revenue and retained new customers.':'Enter annual business revenue and the full one-off value of each project. Project sales count once; churn reduces the existing business revenue baseline. The 24-month goal is revenue in months 13–24.';
      root.querySelectorAll('[data-amount-label]').forEach(el=>el.textContent=monthly?'Monthly value':'Project value');
      q('[data-gp-count]').textContent=`${d.conversions.length} recorded`;
      const leads=d.leads;
      q('[data-gp-source]').textContent=leads&&leads.start===d.measurement_start?`${num(leads.count,0)} leads · ${leads.start} to ${leads.as_of} · Performance Tracker. Ask your AM to refresh when newer leads arrive.`:'Set the tracking start date, then refresh generated leads. The page’s other date filters do not change this planner.';
      const slider=q('[data-gp-slider]');slider.max=Math.max(500,Math.ceil(d.monthly_leads/100)*100);slider.value=d.monthly_leads;
      const result=window.pulseGrowthCalculate(d);
      const salesInput=q('[data-gp-sales]');
      if(document.activeElement!==salesInput)salesInput.value=result.ready?Number((d.monthly_leads*result.rate).toFixed(2)):'';
      salesInput.disabled=!result.ready||busy;
      q('[data-gp-metrics]').replaceChildren();
      if(!form.checkValidity()||!result.ready){q('[data-gp-forecast]').textContent=result.message||'Complete the fields above to see your forecast.';return;}
      q('[data-gp-metrics]').innerHTML=`<div class="gp-metric"><strong>${num(result.leadsPerSale)}</strong>leads per sale</div><div class="gp-metric"><strong>${money(result.average)}</strong>average ${monthly?'monthly':'project'} sale value</div>`+(result.showRate?`<div class="gp-metric"><strong>${num(result.rate*100)}%</strong>lead-to-sale conversion</div>`:'');
      if(result.count!==result.totalCount){const note=document.createElement('p');note.className='gp-note';note.textContent=`Forecast uses ${result.count} conversions through ${leads.as_of}. Later conversions will count after your account manager refreshes leads.`;q('[data-gp-metrics]').appendChild(note);}
      const target=(period,value,goal,needed)=>`<div class="gp-outcome"><h4>${period}</h4><strong>${money(value)}${monthly?'/mo':''}</strong><p>Goal: ${money(goal)}${monthly?'/mo':''}</p><p>${needed==null?'Goal exceeds this planner’s lead range.':`${num(needed,0)} leads/month needed at your observed results.`}</p>${needed!=null?`<button type="button" data-gp-use="${needed}">Try this lead target</button>`:''}</div>`;
      q('[data-gp-forecast]').innerHTML='<div class="gp-outcomes">'+target('12 months',result.at12,d.goal_12,result.needed12)+target('24 months',result.at24,d.goal_24,result.needed24)+'</div>'+chart(result.timeline,d);
      q('[data-gp-forecast]').querySelectorAll('[data-gp-use]').forEach(button=>button.onclick=()=>{field('monthly_leads').value=button.dataset.gpUse;markDirty();render();});
    }
    async function load(){
      message('Loading planner…');q('[data-gp-reload]').hidden=true;
      try{const response=await fetch(root.dataset.endpoint,{credentials:'same-origin',cache:'no-store'});const next=await response.json();if(!response.ok){if(response.status===404&&root.dataset.portal==='true'){root.remove();return;}throw Error(next.error||next.detail||'Could not load the planner.');}apply(next);message('');}
      catch(e){message(e.message,true);q('[data-gp-reload]').hidden=false;}
    }
    async function save(action){
      if(busy)return;if(action!=='deactivate'&&!form.reportValidity())return;
      if(action==='activate'&&(!q('[data-gp-approved]').checked||!q('[data-gp-converted]').checked)){q('[data-gp-dialog-error]').textContent='Confirm both questions before activating.';return;}
      if(action==='deactivate'&&!confirm('Deactivate this planner and remove it from client reports?'))return;
      busy=true;root.querySelectorAll('button').forEach(b=>b.disabled=true);message('Saving…');
      const payload={action,version:state.version,data:action==='deactivate'?state.data:data(),include_in_reports:q('[data-gp-inclusion]').checked,delivery_approved:q('[data-gp-approved]').checked,has_conversion:q('[data-gp-converted]').checked};
      // Avoid accepting edits while the in-flight save owns this version.
      form.querySelectorAll('input,select').forEach(el=>el.disabled=true);
      try{const response=await fetch(root.dataset.endpoint,{method:'POST',credentials:'same-origin',headers:{'Content-Type':'application/json','X-Growth-CSRF':state.csrf},body:JSON.stringify(payload)});const next=await response.json();if(!response.ok){if(response.status===409)q('[data-gp-reload]').hidden=false;throw Error(next.error||next.detail||'Could not save.');}apply(next);dialog.close();message(action==='activate'?'Activated. Publish or re-publish a report to include the planner.':'Saved.');}
      catch(e){message(e.message,true);q('[data-gp-dialog-error]').textContent=e.message;}
      finally{busy=false;root.querySelectorAll('button').forEach(b=>b.disabled=false);form.querySelectorAll('input,select').forEach(el=>el.disabled=false);render();}
    }
    form.addEventListener('input',event=>{
      if(event.target.matches('[data-gp-sales]')){
        const result=window.pulseGrowthCalculate(data());
        if(result.ready&&event.target.value!==''&&event.target.validity.valid)field('monthly_leads').value=Math.ceil(Number(event.target.value)/result.rate-1e-9);
      }
      markDirty();render();
    });
    form.addEventListener('submit',event=>{event.preventDefault();save('save');});
    q('[data-gp-add]').onclick=()=>{row();markDirty();render();};
    q('[data-gp-refresh]').onclick=()=>save('refresh');
    q('[data-gp-slider]').oninput=event=>{field('monthly_leads').value=event.target.value;markDirty();render();};
    q('[data-gp-activate]').onclick=()=>{if(!form.reportValidity())return;q('[data-gp-approved]').checked=false;q('[data-gp-converted]').checked=false;q('[data-gp-dialog-error]').textContent='';dialog.showModal();};
    q('[data-gp-confirm]').onclick=()=>save('activate');q('[data-gp-cancel]').onclick=()=>dialog.close();
    q('[data-gp-deactivate]').onclick=()=>save('deactivate');
    q('[data-gp-reload]').onclick=()=>{if(!dirty||confirm('Discard unsaved changes and load the saved planner?'))load();};
    window.addEventListener('beforeunload',event=>{if(dirty){event.preventDefault();event.returnValue='';}});
    load();
  }
  document.querySelectorAll('[data-growth-planner]').forEach(init);
})();
