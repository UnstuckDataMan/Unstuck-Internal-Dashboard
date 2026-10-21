(function(){
  'use strict';
  function init(root){
    if(root.dataset.ready)return;root.dataset.ready='true';
    const q=s=>root.querySelector(s),form=q('[data-gp-form]'),dialog=q('[data-gp-dialog]');
    let state=null,busy=false,dirty=false,targetBasis='leads',lastActivity=Date.now();
    const fields=['mode','currency','measurement_start','annual_churn','current_revenue','goal_12','goal_24','monthly_leads','annual_sales'];
    const field=n=>form.elements.namedItem(n);
    function message(text,error=false){q('[data-gp-message]').textContent=text;q('[data-gp-message]').dataset.error=String(error);}
    function money(n){return new Intl.NumberFormat('en-GB',{style:'currency',currency:field('currency').value,maximumFractionDigits:0}).format(n);}
    function num(n,d=1){return n.toLocaleString('en-GB',{maximumFractionDigits:d});}
    function markDirty(){lastActivity=Date.now();dirty=true;q('[data-gp-save-status]').textContent='Unsaved changes';}
    function row(conversion={}){
      const el=document.createElement('div');el.className='gp-conversion';el.dataset.id=conversion.id||crypto.randomUUID();
      el.innerHTML='<label>Sale / client <small>optional</small><input data-name maxlength="120"></label><label>Month landed<input data-landed type="month" required></label><label><span data-amount-label>Amount</span><input data-amount type="number" min="0.01" max="1000000000" step="0.01" required></label><label data-months-label>Months paid<input data-months type="number" min="1" max="1200" step="1" required></label><div class="gp-sale-total"><small>Revenue generated</small><strong data-sale-total></strong></div><button type="button" aria-label="Remove conversion">Remove</button>';
      el.querySelector('[data-name]').value=conversion.name||'';
      el.querySelector('[data-landed]').value=(conversion.landed||'').slice(0,7);
      el.querySelector('[data-landed]').max=state.today.slice(0,7);
      el.dataset.originalLanded=conversion.landed||'';
      el.querySelector('[data-months]').value=conversion.months??1;
      el.querySelector('[data-amount]').value=conversion.amount??'';
      el.querySelector('button').onclick=()=>{el.remove();markDirty();render();};
      q('[data-gp-conversions]').appendChild(el);
    }
    function data(){
      const d={};fields.forEach(n=>d[n]=['mode','currency','measurement_start'].includes(n)?field(n).value:Number(field(n).value));
      d.conversions=Array.from(q('[data-gp-conversions]').children).map(el=>({id:el.dataset.id,name:el.querySelector('[data-name]').value,landed:el.dataset.originalLanded.slice(0,7)===el.querySelector('[data-landed]').value?el.dataset.originalLanded:el.querySelector('[data-landed]').value+'-01',amount:Number(el.querySelector('[data-amount]').value),months:Number(el.querySelector('[data-months]').value)}));
      d.leads=state.data.leads;d.target_basis=targetBasis;return d;
    }
    function apply(next){
      state=next;targetBasis=next.data.target_basis||'leads';fields.forEach(n=>field(n).value=next.data[n]??0);field('measurement_start').max=next.today;
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
    function monthLabel(index){
      const date=new Date(Date.UTC(Number(state.today.slice(0,4)),Number(state.today.slice(5,7))-1+index,1));
      return date.toLocaleDateString('en-GB',{month:'short',year:'2-digit',timeZone:'UTC'});
    }
    function revenueChart(values,labels,{bars=false,goals=null,caption='Revenue by month'}={}){
      const max=Math.max(...values,...(goals?goals.map(g=>g.value):[]),1)*1.1;
      const left=88,width=Math.max(672,values.length*30),end=left+width;
      const x=i=>left+(i+.5)*width/values.length,y=v=>180-v/max*140;
      let svg=`<svg class="gp-chart" style="min-width:${end+24}px" viewBox="0 0 ${end+24} 250" role="img" aria-label="${caption}">`;
      for(let i=0;i<=4;i++){
        const value=max*i/4;
        svg+=`<path d="M${left} ${y(value)}H${end}" stroke="#393149"/><text x="${left-10}" y="${y(value)+4}" text-anchor="end">${money(value)}</text>`;
      }
      if(!bars)svg+=`<path d="${values.map((v,i)=>(i?'L':'M')+x(i)+','+y(v)).join(' ')}" fill="none" stroke="#b69aff" stroke-width="3"/>`;
      values.forEach((v,i)=>{
        const title=`${labels[i]}: ${money(v)}`;
        svg+=bars?`<rect x="${x(i)-width/values.length*.32}" y="${y(v)}" width="${width/values.length*.64}" height="${180-y(v)}" rx="3" fill="#b69aff" tabindex="0"><title>${title}</title></rect>`:`<circle cx="${x(i)}" cy="${y(v)}" r="4" fill="#b69aff" tabindex="0"><title>${title}</title></circle>`;
        svg+=`<text transform="translate(${x(i)},198) rotate(-45)" text-anchor="end">${labels[i]}</text>`;
      });
      (goals||[]).forEach(g=>svg+=`<circle cx="${x(g.index)}" cy="${y(g.value)}" r="6" fill="#33d9cd" tabindex="0"><title>${labels[g.index]} goal: ${money(g.value)}</title></circle>`);
      return '<div class="gp-chart-scroll">'+svg+'</svg></div>';
    }
    function chart(values,d){
      return revenueChart(values,values.map((_,i)=>monthLabel(i)),{bars:d.mode==='project',goals:[{index:12,value:d.goal_12},{index:24,value:d.goal_24}],caption:d.mode==='monthly'?'Forecast monthly recurring revenue':'Forecast annual revenue, rolling 12 months'})+
        `<p class="gp-note">Purple: ${d.mode==='monthly'?'forecast monthly recurring revenue':'forecast annual revenue (rolling 12 months)'} · Turquoise: your goals · Scroll to see every month.</p>`;
    }
    function earned(d){
      const host=q('[data-gp-earned]'),monthly=d.mode==='monthly';
      const rows=Array.from(q('[data-gp-conversions]').children);
      rows.forEach(el=>{
        const input=el.querySelector('[data-months]');input.disabled=!monthly||busy;
        el.querySelector('[data-months-label]').hidden=!monthly;
        const landed=el.querySelector('[data-landed]').value;
        input.max=landed?Math.max(1,(Number(state.today.slice(0,4))-Number(landed.slice(0,4)))*12+Number(state.today.slice(5,7))-Number(landed.slice(5,7))+1):1200;
        const amount=Number(el.querySelector('[data-amount]').value),months=monthly?Number(input.value):1;
        el.querySelector('[data-sale-total]').textContent=money(amount*months);
      });
      const recorded=d.conversions;
      if(!recorded.length){host.textContent='Add a conversion to see revenue generated.';return;}
      const monthIndex=s=>Number(s.slice(0,4))*12+Number(s.slice(5,7))-1;
      const now=monthIndex(state.today),start=Math.min(...recorded.map(c=>monthIndex(c.landed)));
      if(!Number.isFinite(start)||now-start>1200||recorded.some(c=>!/^\d{4}-\d{2}-\d{2}$/.test(c.landed)||monthIndex(c.landed)>now||c.amount<=0||!Number.isFinite(c.amount)||(monthly&&(!Number.isInteger(c.months)||c.months<1||monthIndex(c.landed)+c.months-1>now)))){
        host.textContent='Enter a value and valid month for each conversion. Months paid must not extend beyond this month.';return;
      }
      const values=Array(now-start+1).fill(0);
      recorded.forEach(c=>{for(let i=0;i<(monthly?c.months:1);i++)values[monthIndex(c.landed)-start+i]+=c.amount;});
      const total=values.reduce((a,b)=>a+b,0),labels=values.map((_,i)=>monthLabel(start-now+i));
      host.innerHTML=`<h4>Revenue generated from recorded conversions</h4><div class="gp-metric"><strong>${money(total)}</strong>Total recorded revenue</div>`+
        revenueChart(values,labels,{bars:!monthly,caption:'Recorded revenue by month'})+
        `<p class="gp-note">${monthly?'Monthly value × months paid, starting in the month landed. Record paid months only; future revenue appears in the forecast below.':'Each project value is counted once, in the month it landed.'}</p>`;
    }
    function analyse(d,result){
      const host=q('[data-gp-analysis]'),average=q('[data-gp-average-leads]');
      if(!Number.isFinite(result.averageMonthlyLeads)){average.textContent='Refresh generated leads to see the monthly average.';host.textContent='Add your recorded conversions and refresh leads to see how your lead volume compares with your revenue goals.';return;}
      average.innerHTML=`<strong>${num(result.averageMonthlyLeads)}</strong>Average leads generated per month<p class="gp-note">${d.leads.start} to ${d.leads.as_of}. Partial months are prorated.</p>`;
      if(!result.ready){host.textContent='Add a conversion to compare your lead volume with your revenue goals.';return;}
      if(!form.checkValidity()){host.textContent='Complete the fields above to update your analysis.';return;}
      if(!d.goal_12&&!d.goal_24){host.textContent='Enter your 12- and 24-month revenue goals to compare them with your current lead volume.';return;}
      const required=Math.max(result.needed12||0,result.needed24||0),pace=result.averageMonthlyLeads;
      const current=window.pulseGrowthCalculate({...d,monthly_leads:pace});
      let copy=`You currently average ${num(pace)} leads per month, with ${num(result.leadsPerSale)} leads per sale. `;
      if(result.needed12===null||result.needed24===null){
        copy+='The revenue goals exceed the lead range this forecast can calculate. Review the goals, sale value and leads-to-sales performance before setting a lead target.';
      }else if(current.at12+0.01<d.goal_12||current.at24+0.01<d.goal_24){
        copy+=`Current lead volume is below what the revenue goals require. At the recorded sales performance, aim for at least ${num(required,0)} leads per month to meet both goals. `;
        copy+=d.monthly_leads<required?'Your chosen lead target is also below that level. Increase lead volume, or improve lead quality and sales follow-up so fewer leads are needed per sale.':'Your chosen lead target covers that requirement; focus on reaching it consistently while maintaining leads-to-sales performance.';
      }else{
        copy+='Current lead volume is sufficient for both revenue goals under these assumptions. ';
        copy+=d.monthly_leads<required?`Your chosen lead target is lower than the ${num(required,0)} leads per month required, so raise it to keep the plan on track.`:'Maintain lead volume and review new sales and churn regularly to keep the forecast up to date.';
      }
      host.textContent=copy;
    }
    function render(){
      if(!state)return;const d=data(),monthly=d.mode==='monthly';
      q('[data-gp-current-label]').textContent=monthly?'Current monthly revenue':'Current annual revenue';
      q('[data-gp-goal12-label]').textContent=monthly?'Monthly revenue goal in 12 months':'Revenue goal for the next 12 months';
      q('[data-gp-goal24-label]').textContent=monthly?'Monthly revenue goal in 24 months':'Annual revenue goal for year two';
      q('[data-gp-basis-note]').textContent=monthly?'Enter monthly recurring amounts. Annual churn is spread across the year and applies to existing revenue and retained new customers.':'Enter annual business revenue and the full one-off value of each project. Project sales count once; churn reduces the existing business revenue baseline. The 24-month goal is revenue in months 13–24.';
      const symbol={GBP:'£',USD:'$',EUR:'€'}[d.currency];
      root.querySelectorAll('[data-amount-label]').forEach(el=>el.textContent=(monthly?'Monthly value':'Project value')+' ('+symbol+')');
      earned(d);
      q('[data-gp-count]').textContent=`${d.conversions.length} recorded`;
      const leads=d.leads;
      q('[data-gp-source]').textContent=leads&&leads.start===d.measurement_start?`${num(leads.count,0)} leads · ${leads.start} to ${leads.as_of} · Performance Tracker. Ask your AM to refresh when newer leads arrive.`:'Set the tracking start date, then refresh generated leads. The page’s other date filters do not change this planner.';
      const slider=q('[data-gp-slider]');slider.max=Math.max(500,Math.ceil(d.monthly_leads/100)*100);slider.value=d.monthly_leads;
      let result=window.pulseGrowthCalculate(d);
      if(result.ready&&targetBasis==='sales'&&field('annual_sales').value!==''&&field('annual_sales').validity.valid){
        d.monthly_leads=Math.ceil(d.annual_sales/(12*result.rate)-1e-9);
        field('monthly_leads').value=d.monthly_leads;
        slider.max=Math.max(500,Math.ceil(d.monthly_leads/100)*100);slider.value=d.monthly_leads;
        result=window.pulseGrowthCalculate(d);
      }
      const salesInput=q('[data-gp-sales]');
      if(targetBasis==='leads'&&result.ready)salesInput.value=Number((d.monthly_leads*result.rate*12).toFixed(2));
      salesInput.disabled=busy;
      analyse(d,result);
      q('[data-gp-metrics]').replaceChildren();
      if(!form.checkValidity()||!result.ready){q('[data-gp-forecast]').textContent=result.message||'Complete the fields above to see your forecast.';return;}
      q('[data-gp-metrics]').innerHTML=`<div class="gp-metric"><strong>${num(result.leadsPerSale)}</strong>leads per sale</div><div class="gp-metric"><strong>${money(result.average)}</strong>average ${monthly?'monthly':'project'} sale value</div>`+(result.showRate?`<div class="gp-metric"><strong>${num(result.rate*100)}%</strong>lead-to-sale conversion</div>`:'');
      if(result.count!==result.totalCount){const note=document.createElement('p');note.className='gp-note';note.textContent=`Forecast uses ${result.count} conversions through ${leads.as_of}. Later conversions will count after your account manager refreshes leads.`;q('[data-gp-metrics]').appendChild(note);}
      const target=(period,value,goal,needed)=>`<div class="gp-outcome"><h4>${period}</h4><strong>${money(value)}${monthly?'/mo':''}</strong><p>Goal: ${money(goal)}${monthly?'/mo':''}</p><p>${needed==null?'Goal exceeds this planner’s lead range.':`${num(needed,0)} leads/month needed at your observed results.`}</p>${needed!=null?`<button type="button" data-gp-use="${needed}">Try this lead target</button>`:''}</div>`;
      q('[data-gp-forecast]').innerHTML='<div class="gp-outcomes">'+target('12 months',result.at12,d.goal_12,result.needed12)+target('24 months',result.at24,d.goal_24,result.needed24)+'</div>'+chart(result.timeline,d);
      q('[data-gp-forecast]').querySelectorAll('[data-gp-use]').forEach(button=>button.onclick=()=>{targetBasis='leads';field('monthly_leads').value=button.dataset.gpUse;markDirty();render();});
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
    form.addEventListener('input',event=>{
      if(event.target.matches('[data-gp-sales]'))targetBasis='sales';
      if(event.target===field('monthly_leads'))targetBasis='leads';
      markDirty();render();
    });
    form.addEventListener('submit',event=>{event.preventDefault();save('save');});
    q('[data-gp-add]').onclick=()=>{row();markDirty();render();};
    q('[data-gp-refresh]').onclick=()=>save('refresh');
    q('[data-gp-slider]').oninput=event=>{targetBasis='leads';field('monthly_leads').value=event.target.value;markDirty();render();};
    q('[data-gp-activate]').onclick=()=>{if(!form.reportValidity())return;q('[data-gp-approved]').checked=false;q('[data-gp-converted]').checked=false;q('[data-gp-dialog-error]').textContent='';dialog.showModal();};
    q('[data-gp-confirm]').onclick=()=>save('activate');q('[data-gp-cancel]').onclick=()=>dialog.close();
    q('[data-gp-deactivate]').onclick=()=>save('deactivate');
    q('[data-gp-reload]').onclick=()=>{if(!dirty||confirm('Discard unsaved changes and load the saved planner?'))load();};
    window.addEventListener('beforeunload',event=>{if(dirty){event.preventDefault();event.returnValue='';}});
    load();
  }
  document.querySelectorAll('[data-growth-planner]').forEach(init);
})();
