/* Shared state is persisted on the server; the browser submits narrow actions. */
(() => {
  'use strict';
  const $=id=>document.getElementById(id),app=$('gs-app');if(!app)return;
  const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  let board=null,busy=false,submitForm=null,bulkOpen=false,bulkClient=null;
  let processingBatch=null,batchCurrent='',stopBatch=false;
  let selectedClient=new URLSearchParams(location.hash.slice(1)).get('client');
  const stages=['Not started','Started','Purchase made','Setup complete','Failed at QR','Failed at payment','Burned'];
  const closed=stage=>['Setup complete','Burned'].includes(stage);
  const stageOptions=current=>stages.map(stage=>`<option ${stage===current?'selected':''} ${closed(current)&&![current,'Not started','Burned'].includes(stage)?'disabled':''}>${stage}</option>`).join('');
  const state=()=>board?.state||{members:[],devices:[],clients:[],setups:[]};
  const member=id=>state().members.find(m=>m.id===id),device=id=>state().devices.find(d=>d.id===id);
  const name=id=>member(id)?.name||'Unallocated';
  const plannedCount=c=>c&&!c.archived&&!state().setups.some(j=>j.client_id===c.id)?c.planned_domains||0:0;
  const memberOptions=(selected='',pool=false)=>(pool?'<option value="">Unallocated pool</option>':'<option value="">Select team member</option>')+state().members.filter(m=>m.active||m.id===selected).map(m=>`<option value="${esc(m.id)}" ${m.id===selected?'selected':''}>${esc(m.name)}${m.active?'':' (removed)'}</option>`).join('');
  const uses=d=>Math.max(d.use_count||0,d.used_at||d.status==='Used'?1:0);
  const capacity=d=>d.kind==='Phone'?(board?.summary.phone_setups??2):(board?.summary.pc_setups??1);
  const exhausted=d=>uses(d)>=capacity(d);
  const available=(d,jobId)=>{const reserved=state().setups.filter(j=>!j.archived&&['Not started','Started'].includes(j.stage)&&j.id!==jobId&&j.device_id===d.id&&!(d.setup_ids||[]).includes(j.id)).length;return !d.archived&&d.status!=='Blocked'&&uses(d)+reserved<capacity(d);};
  const editing=()=>app.contains(document.activeElement)&&document.activeElement.matches('input,select,textarea');
  const deviceClass=d=>d.status==='Blocked'?'blocked':exhausted(d)?'exhausted '+d.kind.toLowerCase():(uses(d)?'used ':'unused ')+d.kind.toLowerCase();
  const condition=d=>d.status==='Blocked'&&d.block_reason&&d.block_reason!=='Other'?'Block '+d.block_reason:d.status==='Blocked'?'Blocked':`${uses(d)} / ${capacity(d)} ${capacity(d)===1?'setup':'setups'}${exhausted(d)?' · do not use':''}`;
  const conditions=[['Unused','Unused'],['Used','Used'],['Blocked|QR code','Block QR code'],['Blocked|Payment','Block Payment'],['Blocked|Other','Blocked (other)']];
  const inventoryConditions=()=>conditions.flatMap(option=>option[0]==='Used'?Array.from({length:Math.max(2,capacity({kind:'Phone'}),capacity({kind:'PC'}))},(_,i)=>['Used|'+(i+1),i===0?'Used once':i===1?'Used twice':`Used ${i+1} times`]):[option]);
  const conditionData=v=>({status:v.split('|')[0],block_reason:v.startsWith('Blocked|')?v.split('|')[1]:null,...(v.startsWith('Used|')?{use_count:Number(v.split('|')[1])}:{})});
  const label=d=>`${d.kind} ${d.number} · ${name(d.member_id)}`;
  function updateDmarcButton(){const c=state().clients.find(c=>c.id===selectedClient&&!c.archived);for(const id of ['gs-dmarc','gs-client-dns']){$(id).hidden=!c||app.dataset.dnsAdmin!=='true';$(id).disabled=busy||!c||!state().setups.some(j=>j.client_id===c.id&&!j.archived&&j.stage!=='Burned');}}
  function message(text,error=false){$('gs-message').textContent=text;$('gs-message').classList.toggle('error',error);}
  async function api(path='',data){
    const r=await fetch('/api/gmail-setup'+path,{method:data===undefined?'GET':'POST',credentials:'same-origin',cache:'no-store',headers:{'Content-Type':'application/json','X-Domain-CSRF':app.dataset.csrf},...(data===undefined?{}:{body:JSON.stringify(data)})});
    let result;try{result=await r.json();}catch{throw Error('Unexpected server response. Refresh before retrying.');}
    if(!r.ok)throw Error(typeof result.detail==='string'?result.detail:r.status===401?'Sign in again to continue.':'Check the form fields and retry.');
    return result;
  }
  function accept(result){if(board&&result.revision<board.revision)return;board=result;render();}
  async function refresh(){accept(await api());}
  async function run(fn){
    if(busy)return;busy=true;app.setAttribute('aria-busy','true');
    const buttons=[...app.querySelectorAll('button, input, select, textarea'),...$('gs-dialog').querySelectorAll('button')];const disabled=buttons.map(b=>b.disabled);buttons.forEach(b=>b.disabled=true);
    try{await fn();}catch(err){message(err.message,true);if($('gs-dialog').open)$('gs-form-error').textContent=err.message;}
    finally{busy=false;app.removeAttribute('aria-busy');buttons.forEach((b,i)=>{if(b.isConnected)b.disabled=disabled[i];});updateDmarcButton();renderDnsBatch(state().clients.find(c=>c.id===selectedClient));}
  }
  async function change(action,data){const result=await api('/change',{revision:board.revision,action,data});accept(result);return result;}
  function tab(which){for(const view of ['setups','devices']){$('gs-'+view).hidden=view!==which;$('gs-tab-'+view).setAttribute('aria-selected',String(view===which));}}
  function render(){
    const s=board.summary,st=state();
    $('gs-stats').innerHTML=[['not_started','Not started'],['started','Started'],['purchased','Purchase made'],['complete','Setup complete']].map(([key,title])=>`<div class="gs-stat"><strong>${s[key]}</strong><span>${title}</span></div>`).join('')+`<div class="gs-stat gs-reserve-stat"><span>Setup capacity / required</span><div class="gs-reserve-ratio"><strong>${s.available_setups}</strong><span>/</span><strong id="gs-reserve-target">${s.target}</strong></div><b class="gs-shortfall ${s.low_stock?'low':'ready'}">${s.low_stock?`Order ${s.shortfall} more ${s.order_kind==='Phone'?'phones':'laptops / PCs'}`:'Setup capacity covered'}</b><p class="gs-reserve-calculation">Required: ${s.todo} to do${s.planned_todo?` (${s.actual_todo} actual + ${s.planned_todo} planned)`:""} + ${s.contingency} contingency = ${s.target} setups.<br>Available: ${s.phone_available} phone + ${s.pc_available} laptop / PC${s.purchased_setups?` + ${s.purchased_setups} already purchased`:""} = ${s.available_setups} setups.<br>${s.available_setups>s.target?`Surplus: ${s.available_setups} − ${s.target} = ${s.available_setups-s.target} setup slots.`:`Gap: ${s.target} − ${s.available_setups} = ${s.setup_shortfall} setups.`}<br>${s.setup_shortfall?`${s.setup_shortfall} ÷ ${s.order_capacity}, rounded up = ${s.shortfall} new ${s.order_kind==='Phone'?'phones':'laptops / PCs'} (${s.shortfall*s.order_capacity} extra setup slots).`:`No new devices needed. New ${s.order_kind==='Phone'?'phones':'laptops / PCs'} cover ${s.order_capacity} setups each.`}</p></div>`;
    $('gs-phone-capacity').value=s.phone_setups;$('gs-pc-capacity').value=s.pc_setups;$('gs-order-kind').value=s.order_kind;
    $('gs-device-rules').textContent=`Phones: ${s.phone_setups} setups each · Laptops / PCs: ${s.pc_setups} setups each.`;
    $('gs-stock').hidden=!s.low_stock;
    $('gs-stock').innerHTML=`<div><strong>Order ${s.shortfall} more ${s.order_kind==='Phone'?'phones':'laptops / PCs'}</strong><small>${s.available_setups} setup capacity / ${s.target} required · ${s.todo} to do${s.planned_todo?` including ${s.planned_todo} planned`:""} + ${s.contingency} contingency setups. ${s.setup_shortfall} setups short; ${s.order_capacity} per new ${s.order_kind==='Phone'?'phone':'laptop / PC'}, rounded up.</small></div><button data-stock-add>Add inventory</button>`;
    const chosen=$('gs-filter-manager').value;$('gs-filter-manager').innerHTML='<option value="">Everyone</option>'+st.members.map(m=>`<option value="${esc(m.id)}">${esc(m.name)}${m.active?'':' (removed)'}</option>`).join('');$('gs-filter-manager').value=chosen;
    renderSetups();renderTeam();$('gs-updated').textContent=board.updated_at?`Updated ${new Date(board.updated_at).toLocaleString('en-GB')}`:'Add your actual device inventory to get started';
  }
  function openClient(id){
    selectedClient=id||null;
    history.pushState(null,'',location.pathname+location.search+(id?'#client='+encodeURIComponent(id):''));
    renderSetups();
  }
  function renderDomain(s){
    const d=device(s.device_id),p=s.purchase_device,sync=s.sync||{},kind=s.section==='client'?'Client':'Biz Dev';
    const failure=s.stage.startsWith('Failed at'),published=sync.status==='synced',burned=s.stage==='Burned';
    const cardStage=burned||failure?' gs-domain-danger':s.stage==='Setup complete'?' gs-domain-complete':['Started','Purchase made'].includes(s.stage)?' gs-domain-started':'';
    const dns=s.dmarc&&s.dmarc.email===s.email&&s.dmarc.domain===s.domain?s.dmarc:null;
    const setup=s.dns_setup?.domain===s.domain&&s.dns_setup?.email===s.email?s.dns_setup:null;
    const setupReady=setup?.status==='verified'&&setup.configuration_version===2;
    const ownerLocked=burned||published||s.stage==='Setup complete'||Boolean(p&&s.stage==='Purchase made'),deviceLocked=burned||Boolean(d&&(published||p&&['Purchase made','Setup complete'].includes(s.stage)));
    const options='<option value="">Device not added yet</option>'+state().devices.filter(x=>x.id===s.device_id||available(x,s.id)&&x.member_id===s.operator_id).map(x=>`<option value="${esc(x.id)}" ${x.id===s.device_id?'selected':''}>${esc(label(x)+' · '+condition(x))}</option>`).join('');
    const safeLink=sync.url&&/^https:\/\/docs\.google\.com\/spreadsheets\/d\/[A-Za-z0-9_-]+\/edit#gid=\d+&range=A\d+$/.test(sync.url)?sync.url:null;
    return `<article class="gs-domain${cardStage}" data-setup="${esc(s.id)}"><div class="gs-domain-head"><div><h4>${esc(s.domain)}</h4>${s.order&&s.stage==='Setup complete'?`<small class="gs-order">Sheet #${s.order}</small>`:''}</div><div class="gs-domain-buttons">${!published&&!burned?`<button class="gs-small" data-edit-setup="${esc(s.id)}">Details</button>`:''}<button class="gs-small gs-danger" data-remove-setup="${esc(s.id)}">Delete</button></div></div>
      <div class="gs-domain-controls"><label>Email<input type="email" data-set-email="${esc(s.id)}" value="${esc(s.email)}" placeholder="Add sending email" maxlength="254" ${burned||published||s.stage==='Setup complete'?'disabled':''}></label><label>Setup person<select data-set-operator="${esc(s.id)}" ${ownerLocked?'disabled':''}>${memberOptions(s.operator_id)}</select></label><label class="gs-device-choice">${d?`<strong class="gs-phone-number ${deviceClass(d)}">${esc((p?.id===d.id?p.label:d.kind+' '+d.number))}</strong>`:'Device'}<select data-set-device="${esc(s.id)}" aria-label="Device for ${esc(s.domain)}" ${deviceLocked?'disabled':''}>${options}</select></label><label class="gs-stage-control">Stage<select data-set-stage="${esc(s.id)}" aria-label="Stage for ${esc(s.domain)}">${stageOptions(s.stage,Boolean(s.purchase_at))}</select></label></div>
      ${burned||!d||failure||d?.status==='Blocked'||d&&exhausted(d)?`<div class="gs-domain-meta">${burned?'<span class="gs-badge burned">Burned</span>':''}${!d?'<span class="gs-badge purchased">Device not added yet</span>':''}${failure?`<span class="gs-badge blocked">${esc(s.stage)}</span>`:''}${d?.status==='Blocked'?`<span class="gs-badge blocked">${esc(condition(d))} · do not use</span>`:d&&exhausted(d)?`<span class="gs-badge exhausted">${esc(condition(d))}</span>`:''}</div>`:''}
      ${!burned&&(app.dataset.dnsAdmin==='true'||dns||setup)?`<div class="gs-domain-dmarc">${dns?`<span class="gs-badge ${dns.status==='verified'?'complete':'blocked'}" title="${esc(dns.message)}">DMARC ${dns.status==='verified'?'verified':'needs attention'}</span>`:''}${app.dataset.dnsAdmin==='true'?`<button class="gs-small" data-setup-dns="${esc(s.id)}">${setupReady?'Recheck setup':'One-click setup'}</button>`:''}${setup?`<span class="gs-badge ${setupReady?'complete':'blocked'}" title="${esc(setup.message)}">${setupReady?'DNS ready':setup.status==='verified'?'Google mail setup needed':'DNS setup needs attention'}</span>`:''}</div>`:''}
      <details class="gs-domain-details"><summary>Account details</summary><dl class="gs-facts"><div><dt>Account manager</dt><dd>${esc(name(s.account_manager_id))}</dd></div><div><dt>Purchase asset</dt><dd>${esc(p?.asset_tag||d?.asset_tag||'Not allocated')}</dd></div><div><dt>Sheet destination</dt><dd>${kind}</dd></div></dl>${(s.device_failures||[]).length?`<p class="gs-help">${s.device_failures.map(f=>esc(`${f.stage} · ${f.device.label?f.device.label+' · '+f.device.member+' · '+f.device.asset_tag:'Device not recorded'}`)).join('<br>')}</p>`:''}${s.notes?`<p class="gs-help">${esc(s.notes)}</p>`:''}</details><div class="gs-sync ${sync.status==='error'?'error':''}">${published?`${s.stage==='Setup complete'?'Live in':'Previously synced in'} ${esc(sync.tab)} · row ${sync.row}${safeLink?` · <a href="${esc(safeLink)}" target="_blank" rel="noopener noreferrer">Open sheet row ↗</a>`:''}`:sync.status==='error'&&!burned?`Sheet sync needs attention: ${esc(sync.message)}${safeLink?` · <a href="${esc(safeLink)}" target="_blank" rel="noopener noreferrer">Open affected sheet row ↗</a>`:''}`:s.stage==='Setup complete'?'Saved as complete · sheet sync pending':''}</div>${s.stage==='Setup complete'&&!published?`<button class="gs-small gs-retry" data-sync="${esc(s.id)}">Retry sheet sync</button>`:''}</article>`;
  }
  function renderSetups(){
    const search=$('gs-search').value.toLowerCase(),stage=$('gs-filter-stage').value,manager=$('gs-filter-manager').value,section=$('gs-filter-section').value;
    const all=state().setups.filter(s=>!s.archived);
    const matches=s=>(!stage||s.stage===stage)&&(!manager||s.account_manager_id===manager)&&(!section||s.section===section)&&(!search||[s.domain,s.email,s.client_name].join(' ').toLowerCase().includes(search));
    const clients=(state().clients||[]).filter(c=>!c.archived);
    const current=clients.find(c=>c.id===selectedClient);
    if(!current)selectedClient=null;
    $('gs-client-list').hidden=Boolean(current);$('gs-client-detail').hidden=!current;
    $('gs-back-clients').hidden=!current;
    $('gs-client-heading').textContent=current?current.name:'Clients';
    $('gs-client-caption').textContent=current?`${name(current.account_manager_id)} · ${current.section==='client'?'Client':'Biz Dev'} · ${plannedCount(current)?`${plannedCount(current)} planned domains included in the forecast`:'Completed domains move below in-progress work; burned domains appear last'}`:'Choose a client to view its domains and setup stages.';
    $('gs-add-setup').textContent=current?'+ Add domains':'+ Client';
    $('gs-edit-client').hidden=!current;$('gs-delete-client').hidden=!current;
    updateDmarcButton();renderDmarc(current);renderDnsBatch(current);
    if(current&&bulkClient!==current.id){bulkClient=current.id;bulkOpen=!all.some(s=>s.client_id===current.id);$('gs-bulk-domains').value='';$('gs-bulk-sender').value='';$('gs-bulk-operator').innerHTML=memberOptions(member(current.operator_id)?.active?current.operator_id:'');}
    $('gs-bulk-panel').hidden=!current||!bulkOpen;
    const visible=clients.slice().sort((a,b)=>a.name.localeCompare(b.name)).filter(c=>{
      const domains=all.filter(s=>s.client_id===c.id);
      return domains.length?domains.some(matches):(!stage||stage==='Not started'&&plannedCount(c)>0)&&(!manager||c.account_manager_id===manager)&&(!section||c.section===section)&&(!search||c.name.toLowerCase().includes(search));
    });
    $('gs-client-list').innerHTML=visible.length?visible.map(c=>{
      const domains=all.filter(s=>s.client_id===c.id),complete=domains.filter(s=>s.stage==='Setup complete').length,burned=domains.filter(s=>s.stage==='Burned').length,planned=plannedCount(c);
      return `<button class="gs-client-card" data-open-client="${esc(c.id)}"><div><strong>${esc(c.name)}</strong><small>${esc(name(c.account_manager_id))} · ${c.section==='client'?'Client':'Biz Dev'}</small></div><div class="gs-client-counts"><span>${planned?`${planned} planned domains`:`${domains.length} ${domains.length===1?'domain':'domains'}`}</span><span>${planned||domains.length-complete-burned} to do</span><span class="gs-complete-count">${complete} complete</span>${burned?`<span class="gs-burned-count">${burned} burned</span>`:''}<span aria-hidden="true">→</span></div></button>`;
    }).join(''):'<div class="gs-empty"><strong>No clients to show</strong>Add a client first, then add its domains.</div>';
    if(current){
      const domains=all.filter(s=>s.client_id===current.id&&matches(s));
      const pending=domains.filter(s=>!closed(s.stage)),complete=domains.filter(s=>s.stage==='Setup complete'),burned=domains.filter(s=>s.stage==='Burned');
      const order=(a,b)=>(a.order||0)-(b.order||0)||a.domain.localeCompare(b.domain);
      const pendingOrder=(a,b)=>(a.created_at||'').localeCompare(b.created_at||'');
      $('gs-setup-list').innerHTML=domains.length?`${pending.length?`<h4 class="gs-list-heading">In progress <span>${pending.length}</span></h4>${pending.sort(pendingOrder).map(renderDomain).join('')}`:''}${complete.length?`<h4 class="gs-list-heading gs-complete-heading">Complete <span>${complete.length}</span></h4>${complete.sort(order).map(renderDomain).join('')}`:''}${burned.length?`<h4 class="gs-list-heading gs-burned-heading">Burned <span>${burned.length}</span></h4>${burned.sort((a,b)=>(a.burned_at||'').localeCompare(b.burned_at||'')).map(renderDomain).join('')}`:''}`:`<div class="gs-empty"><strong>No domains to show</strong>${plannedCount(current)?`${plannedCount(current)} planned domains are included in the forecast. Add actual domains to replace this placeholder.`:'Add a domain to this client or change your filters.'}</div>`;
    }else $('gs-setup-list').innerHTML='';
  }
  function renderDmarc(client){
    const batch=client?.dmarc_batch,el=$('gs-dmarc-results');el.hidden=!batch;
    if(!batch){el.innerHTML='';return;}
    const verified=Object.values(batch.results).filter(r=>r.status==='verified').length;
    el.innerHTML=`<details class="gs-card gs-dmarc-results" open><summary>Latest DMARC run: ${verified} / ${batch.items.length} verified</summary><ul>${batch.items.map(item=>{const r=batch.results[item.id]||{},ok=r.status==='verified';return `<li><strong>${esc(item.domain)}</strong><span class="gs-badge ${ok?'complete':r.status==='needs review'?'blocked':'purchased'}">${ok?'Verified':r.status==='running'?'Configuring':r.status==='needs review'?'Needs attention':'Pending'}</span><small>${esc(item.email?'Reports: '+item.email:'Sending email missing')}${r.message?' · '+esc(r.message):''}</small></li>`;}).join('')}</ul></details>`;
  }
  function setupDmarc(jobId){
    const client=state().clients.find(c=>c.id===selectedClient);if(!client)return;
    const id=client.id,items=state().setups.filter(j=>j.client_id===id&&!j.archived&&j.stage!=='Burned'&&(!jobId||j.id===jobId));if(!items.length)return;
    dialog('Set up DMARC for '+(jobId?items[0].domain:client.name),`<p class="gs-dmarc-warning"><strong>Only proceed if you’ve completed purchase.</strong></p><p>This adds DMARC to ${items.length===1?'this domain':'these '+items.length+' domains'} in Porkbun. Both aggregate and failure reports go to each domain’s sending email.</p><ul class="gs-dmarc-preview">${items.map(j=>`<li><strong>${esc(j.domain)}</strong><small>${esc(j.email||'Sending email missing — will be skipped')}</small></li>`).join('')}</ul><p><code>v=DMARC1; p=none; pct=100; rua=mailto:[sending email]; ruf=mailto:[sending email]; fo=1</code></p><p>TXT host: <strong>_dmarc</strong> · TTL: <strong>600</strong>. Existing matching records are checked; conflicting records are left for review.</p>`,async()=>{
      accept(await api('/clients/'+id+'/dmarc/start',{revision:board.revision,confirmed_purchased:true,setup_ids:items.map(j=>j.id)}));$('gs-dialog').close();
      const batch=state().clients.find(c=>c.id===id).dmarc_batch;
      for(const [index,item] of batch.items.entries()){
        if(!item.email||batch.results[item.id]?.status==='needs review')continue;
        message(`Configuring DMARC ${index+1} / ${batch.items.length}: ${item.domain}…`);
        accept(await api('/clients/'+id+'/dmarc/'+batch.id+'/'+item.id,{revision:board.revision}));
      }
      const finished=state().clients.find(c=>c.id===id).dmarc_batch,verified=Object.values(finished.results).filter(r=>r.status==='verified').length;
      message(`DMARC: ${verified} verified · ${finished.items.length-verified} need attention.`,verified!==finished.items.length);
    });$('gs-dialog-save').textContent='Proceed';
  }
  function renderDnsBatch(client){
    const batch=client?.dns_setup_batch,el=$('gs-dns-batch');el.hidden=!batch;
    if(!batch){el.innerHTML='';return;}
    const results=batch.results||{},ready=batch.items.filter(i=>results[i.id]?.status==='verified').length,attention=batch.items.filter(i=>results[i.id]?.status==='needs review').length,done=ready+attention,total=batch.items.length;
    const active=processingBatch===client.id,pending=done<total;
    const status=active?(stopBatch?'Pausing after the current domain…':`Setting up ${batchCurrent||'next domain'}…`):pending?(batch.paused_reason||'Paused. Resume to continue the waiting domains.'):'Finished';
    el.innerHTML=`<section class="gs-card gs-batch-card"><div class="gs-batch-header"><h4>All-domain setup</h4><strong>${done} / ${total} processed</strong></div><progress max="${total}" value="${done}" aria-label="Domain setup progress"></progress><p class="gs-batch-counts"><span>${ready} ready</span><span>${attention} need attention</span><span>${total-done} waiting</span></p><p class="gs-batch-status" role="status">${esc(status)}</p>${active?'<p class="gs-help">Keep this page open while setup runs. You can resume after a refresh.</p>':''}<div class="gs-actions">${app.dataset.dnsAdmin==='true'?(active?`<button data-pause-dns ${stopBatch?'disabled':''}>${stopBatch?'Pausing…':'Pause after this domain'}</button>`:pending?'<button data-resume-dns class="gs-primary">Resume setup</button>':''):''}</div><details><summary>Domain results</summary><ul class="gs-batch-results">${batch.items.map(i=>{const r=results[i.id]||{},ok=r.status==='verified',failed=r.status==='needs review';return `<li><strong>${esc(i.domain)}</strong><span class="gs-badge ${ok?'complete':failed?'blocked':'purchased'}">${ok?'Ready':failed?'Needs attention':active&&batchCurrent===i.domain||r.status==='running'?'In progress':'Waiting'}</span>${r.message?`<small>${esc(r.message)}</small>`:''}</li>`;}).join('')}</ul></details></section>`;
  }
  async function processDnsBatch(clientId,batchId){
    processingBatch=clientId;stopBatch=false;
    try{
      while(!stopBatch){
        const client=state().clients.find(c=>c.id===clientId),batch=client?.dns_setup_batch;
        if(!batch||batch.id!==batchId)throw Error('The batch changed. Refresh to see its results.');
        const next=batch.items.find(i=>!['verified','needs review'].includes(batch.results[i.id]?.status));
        if(!next)break;
        batchCurrent=next.domain;renderDnsBatch(client);
        accept(await api('/clients/'+clientId+'/dns-setup/'+batchId+'/'+next.id,{revision:board.revision}));
        if(state().clients.find(c=>c.id===clientId).dns_setup_batch.paused_reason)break;
      }
      const batch=state().clients.find(c=>c.id===clientId).dns_setup_batch,ready=Object.values(batch.results).filter(r=>r.status==='verified').length,attention=Object.values(batch.results).filter(r=>r.status==='needs review').length;
      message(`${ready} / ${batch.items.length} domains ready${attention?` · ${attention} need attention`:''}${ready+attention<batch.items.length?' · setup paused; resume when ready':''}.`,attention>0);
    }catch(err){
      // A lost response may still have completed on the server. Read saved
      // results before offering Resume; do not blindly repeat an active write.
      try{await refresh();}catch{}
      throw err;
    }finally{processingBatch=null;batchCurrent='';renderDnsBatch(state().clients.find(c=>c.id===selectedClient));}
  }
  function setupAllDomains(){
    const client=state().clients.find(c=>c.id===selectedClient);if(!client)return;
    const items=state().setups.filter(j=>j.client_id===client.id&&!j.archived&&j.stage!=='Burned');if(!items.length)return;
    dialog('Set up all domains for '+client.name,`<p class="gs-dmarc-warning"><strong>Only proceed if you’ve completed purchase for all ${items.length} domains.</strong></p><label>Client website URL<input id="gs-batch-destination" inputmode="url" placeholder="https://clientwebsite.com" value="${esc(client.website_url||'')}" maxlength="2048" required></label><p>Set up all ${items.length} active domains for this client: remove default DNS, add Google MX/SPF, tracking CNAME and DMARC, then replace root forwarding with the URL above.</p><p>Domains run one at a time. Burned and deleted domains are excluded. Keep this page open; saved progress can be resumed after a refresh.</p><details><summary>View ${items.length} domains</summary><ul class="gs-dmarc-preview">${items.map(i=>`<li><strong>${esc(i.domain)}</strong><small>${esc(i.email||'Sending email missing — needs attention')}</small></li>`).join('')}</ul></details>`,async()=>{
      const result=await api('/clients/'+client.id+'/dns-setup/start',{revision:board.revision,confirmed_purchased:true,destination:value('gs-batch-destination')});
      accept(result);$('gs-dialog').close();await processDnsBatch(client.id,state().clients.find(c=>c.id===client.id).dns_setup_batch.id);
    });$('gs-dialog-save').textContent='Proceed';
  }
  $('gs-dns-batch').onclick=e=>{
    if(e.target.closest('[data-pause-dns]')){stopBatch=true;renderDnsBatch(state().clients.find(c=>c.id===selectedClient));}
    if(e.target.closest('[data-resume-dns]')){const c=state().clients.find(c=>c.id===selectedClient);if(c)run(()=>processDnsBatch(c.id,c.dns_setup_batch.id));}
  };
  function previewDnsSetup(jobId){
    return run(async()=>{
      message('Checking the default DNS records for this domain…');
      accept(await api('/setups/'+jobId+'/dns-setup/preview',{revision:board.revision}));
      const job=state().setups.find(j=>j.id===jobId),preview=job.dns_setup_preview,client=state().clients.find(c=>c.id===job.client_id);
      dialog('One-click setup for '+job.domain,`<p class="gs-dmarc-warning"><strong>Only proceed if you’ve completed purchase.</strong></p>
        <label>Client website URL<input id="gs-dns-destination" type="text" inputmode="url" placeholder="https://clientwebsite.com" value="${esc(client.website_url||job.dns_setup?.destination||'')}" maxlength="2048" required></label>
        <p>Proceed will set up <strong>${esc(job.domain)}</strong> in Porkbun:</p>
        <ul class="gs-dmarc-preview"><li><strong>Remove ${preview.records.length} default DNS records</strong><small>Porkbun forwarding MX, default SPF and ACME challenges.</small></li><li><strong>Google email records</strong><small>MX smtp.google.com · priority 1 · TTL 600<br>TXT v=spf1 include:_spf.google.com ~all · TTL 600</small></li><li><strong>Tracking CNAME</strong><small>t.${esc(job.domain)} → open.sleadtrack.com · TTL 600</small></li><li><strong>DMARC</strong><small>Reports to ${esc(job.email)} · TTL 600</small></li><li><strong>Website forwarding</strong><small>Replace existing root forwarding with the URL above · 302 redirect · wildcard enabled</small></li></ul>
        <details><summary>View DNS records</summary><ul class="gs-dmarc-preview gs-cleanup-preview">${preview.records.map(r=>`<li><strong>${esc(r.type)} · ${esc(r.name)}</strong><small>${esc(r.content)}</small></li>`).join('')}</ul>${preview.records.length?'':'<p>No default records to remove.</p>'}${preview.forwarding_records.length?`<p>Root forwarding to replace:</p><ul class="gs-dmarc-preview">${preview.forwarding_records.map(r=>`<li>${esc(r.location)}</li>`).join('')}</ul>`:''}<p><code>${esc(preview.additions[1].payload.content)}</code></p></details>
        <p>Other mail and verification records are kept.</p>`,async()=>{
        message('Setting up DNS, tracking, DMARC and website forwarding…');
        const result=await api('/setups/'+jobId+'/dns-setup/apply',{revision:board.revision,preview_id:preview.id,confirmed_purchased:true,destination:value('gs-dns-destination')});
        accept(result);const cleanup=result.state.setups.find(j=>j.id===jobId).dns_setup;
        message(cleanup.message,cleanup.status!=='verified');
      });$('gs-dialog-save').textContent='Proceed';message('Review the records before proceeding.');
    });
  }
  function deviceDomains(d){
    return state().setups.filter(j=>j.device_id===d.id||(d.setup_ids||[]).includes(j.id)||j.purchase_device?.id===d.id||(j.purchase_device_history||[]).some(p=>p.id===d.id)||(j.device_failures||[]).some(f=>f.device?.id===d.id)).map(j=>{
      const failure=(j.device_failures||[]).filter(f=>f.device?.id===d.id&&!f.reverted_at).at(-1);
      const reset=Boolean(j.device_usage_resets?.[d.id]);
      const status=reset?'Reset to Not started · not counted':failure?failure.stage:(d.setup_ids||[]).includes(j.id)?`Used · ${j.stage}`:j.device_id===d.id?`Allocated · ${j.stage}`:`Previous device · ${j.stage}`;
      return {job:j,status,reset};
    }).sort((a,b)=>Number(a.reset)-Number(b.reset)||Number(a.job.archived)-Number(b.job.archived));
  }
  function deviceDomainPreview(d){const rows=deviceDomains(d);return rows.length?`<small class="gs-device-domains">${rows.slice(0,2).map(r=>esc(r.job.domain)+(r.reset?' · reset':'')).join('<br>')}${rows.length>2?`<br>+${rows.length-2} more`:''}</small>`:'';}
  function deviceDomainDetails(d){const rows=deviceDomains(d);return `<section class="gs-device-history"><h4>Domains linked to this device</h4>${rows.length?`<ul>${rows.map(({job,status})=>`<li>${!job.archived&&state().clients.some(c=>c.id===job.client_id&&!c.archived)?`<button type="button" data-device-domain="${esc(job.id)}">${esc(job.domain)} ↗</button>`:`<strong>${esc(job.domain)}</strong>`}<small>${esc(job.client_name)} · ${esc(status)}${job.archived?' · deleted from tracker':''}</small></li>`).join('')}</ul>`:'<p>No domains linked yet.</p>'}</section>`;}
  function renderTeam(){
    const active=state().members.filter(m=>m.active),pool={id:null,name:'Unallocated pool',email:''};
    $('gs-team-list').innerHTML=[...active,pool].map(m=>{
      const ds=state().devices.filter(d=>!d.archived&&d.member_id===m.id).sort((a,b)=>a.kind.localeCompare(b.kind)||a.number-b.number);
      return `<article class="gs-card" data-member="${esc(m.id||'pool')}"><div class="gs-card-head"><div><h4>${esc(m.name)}</h4><small>${esc(m.email)}</small></div><div class="gs-actions">${m.id?`<button class="gs-small" data-edit-member="${esc(m.id)}">Edit</button><button class="gs-small gs-danger" data-remove-member="${esc(m.id)}">Remove</button>`:''}<button class="gs-small" data-add-for="${esc(m.id||'')}">+ Devices</button></div></div><div class="gs-device-counts"><span><strong>${ds.filter(d=>d.kind==='Phone').length}</strong> phones${m.id?' / 10':''}</span><span><strong>${ds.filter(d=>d.kind==='PC').length}</strong> PCs${m.id?' / 10':''}</span><span><strong>${ds.filter(d=>d.status==='Unused'&&!uses(d)).length}</strong> unused</span><span><strong>${ds.filter(d=>d.status==='Blocked'||exhausted(d)).length}</strong> unavailable</span></div>${ds.length?`<div class="gs-device-grid">${ds.map(d=>`<button class="gs-device ${deviceClass(d)}" data-device="${esc(d.id)}" aria-label="${esc(label(d)+' · '+condition(d)+' · '+d.asset_tag)}"><strong class="gs-device-number">${esc(d.kind+' '+d.number)}</strong><span>${esc(condition(d))}</span><small>${esc(d.asset_tag)}</small>${deviceDomainPreview(d)}</button>`).join('')}</div>`:'<p class="gs-help">No devices allocated.</p>'}</article>`;
    }).join('');
  }
  function dialog(title,html,onSubmit){$('gs-dialog-save').textContent='Save';$('gs-dialog-title').textContent=title;$('gs-form-fields').innerHTML=html;$('gs-form-error').textContent='';submitForm=onSubmit;$('gs-dialog').showModal();}
  const value=id=>$(id).value.trim();
  const field=(title,id,val='',extra='')=>`<label>${title}<input id="${id}" value="${esc(val)}" ${extra}></label>`;
  function editMember(id){const m=member(id)||{};dialog(id?'Edit team member':'Add team member',`<div class="gs-form-grid">${field('Name','gs-member-name',m.name,'required maxlength="100"')}${field('Team email (optional)','gs-member-email',m.email,'type="email" maxlength="254"')}</div>`,async()=>{await change('member_save',{...(id?{id}:{}),name:value('gs-member-name'),email:value('gs-member-email')});message('Team member saved.');});}
  function addDevices(owner=''){
    const counts=Array.from({length:11},(_,i)=>`<option value="${i}">${i}</option>`).join('');
    dialog('Add device inventory',`<div class="gs-form-grid"><label class="gs-wide">Allocate to<select id="gs-device-owner">${memberOptions(owner,true)}</select></label><label>Number of phones<select id="gs-device-phones">${counts}</select></label><label>Number of PCs<select id="gs-device-pcs">${counts}</select></label><label class="gs-wide">Current condition<select id="gs-device-condition">${inventoryConditions().map(([v,l])=>`<option value="${esc(v)}">${l}</option>`).join('')}</select></label></div>`,async()=>{await change('devices_add',{member_id:value('gs-device-owner')||null,phones:Number(value('gs-device-phones')),pcs:Number(value('gs-device-pcs')),...conditionData(value('gs-device-condition'))});message('Device inventory added.');});
  }
  function editDevice(id){const d=device(id);dialog(label(d),`<div class="gs-form-grid">${field('Physical asset label','gs-asset-tag',d.asset_tag,'required maxlength="80"')}<label>Allocate to<select id="gs-transfer-owner">${memberOptions(d.member_id,true)}</select></label><label>Condition<select id="gs-asset-status">${conditions.map(([v,l])=>`<option value="${esc(v)}" ${v===(d.status==='Blocked'?'Blocked|'+(d.block_reason||'Other'):d.status)?'selected':''} ${d.status==='Blocked'&&!v.startsWith('Blocked')||d.used_at&&v==='Unused'?'disabled':''}>${l}</option>`).join('')}</select></label></div>${deviceDomainDetails(d)}<button type="button" id="gs-remove-device" class="gs-danger">Remove from inventory</button>`,async()=>{await change('device_update',{id,member_id:value('gs-transfer-owner')||null,...conditionData(value('gs-asset-status')),asset_tag:value('gs-asset-tag')});message('Device updated.');});for(const b of $('gs-form-fields').querySelectorAll('[data-device-domain]'))b.onclick=()=>{const job=state().setups.find(j=>j.id===b.dataset.deviceDomain);$('gs-dialog').close();tab('setups');$('gs-search').value='';$('gs-filter-stage').value='';$('gs-filter-manager').value='';$('gs-filter-section').value='';openClient(job.client_id);const row=$('gs-setup-list').querySelector(`[data-setup="${job.id}"]`);row?.scrollIntoView({block:'center',behavior:'smooth'});row?.classList.add('gs-domain-highlight');setTimeout(()=>row?.classList.remove('gs-domain-highlight'),2500);};$('gs-remove-device').onclick=()=>{if(confirm('Remove this device from active inventory? Purchase history will be retained.'))run(async()=>{await change('device_remove',{id});$('gs-dialog').close();message('Device removed from active inventory.');});};}
  function addClient(id){
    const c=state().clients.find(c=>c.id===id)||{};
    const canPlan=!state().setups.some(j=>j.client_id===c.id);
    dialog(id?'Edit client':'Add client',`<div class="gs-form-grid">${field('Client / company name','gs-client-name',c.name,'required maxlength="150"')}<label>Section<select id="gs-client-section" ${id?'disabled':''}><option value="client" ${c.section==='client'?'selected':''}>Client · 1) Emails - entry</option><option value="bizdev" ${c.section==='bizdev'?'selected':''}>Biz Dev · 1) BD Emails - entry</option></select></label><label>Account manager<select id="gs-client-manager" required>${memberOptions(c.account_manager_id)}</select></label><label>Default setup person (optional)<select id="gs-client-operator">${memberOptions(member(c.operator_id)?.active?c.operator_id:'')}</select></label>${canPlan?`<label class="gs-wide">Planned domains (optional)<input id="gs-client-planned" type="number" inputmode="numeric" min="0" max="5000" step="1" value="${c.planned_domains||0}"><small>Used until you add actual domains.</small></label>`:''}</div>`,async()=>{
      const existing=new Set(state().clients.map(c=>c.id));
      const r=await change('client_save',{...(id?{id}:{}),name:value('gs-client-name'),section:value('gs-client-section'),account_manager_id:value('gs-client-manager'),operator_id:value('gs-client-operator')||null,planned_domains:canPlan?Number(value('gs-client-planned')||0):0});
      if(!id)openClient(r.state.clients.find(c=>!existing.has(c.id)).id);message(id?'Client updated.':'Client added.');
    });
  }
  function editSetup(id){
    const s=state().setups.find(s=>s.id===id)||{};
    const c=state().clients.find(c=>c.id===(s.client_id||selectedClient));
    if(!c)return addClient();
    dialog(id?'Edit domain':`Add domain to ${c.name}`,`<div class="gs-form-grid"><label>Setup owner<select id="gs-job-operator">${memberOptions(s.operator_id||c.operator_id)}</select></label>${field('Email','gs-job-email',s.email,'type="email" maxlength="254"')}<label>Stage<select id="gs-job-stage">${stageOptions(s.stage||'Not started',Boolean(s.purchase_at))}</select></label><label class="gs-wide">Device (optional)<select id="gs-job-device"><option value="">Allocate a device later</option></select></label><label class="gs-wide">Notes (optional)<textarea id="gs-job-notes" maxlength="1000">${esc(s.notes)}</textarea></label></div>`,async()=>{
      const r=await change('setup_save',{...(id?{id}:{}),client_id:c.id,stage:value('gs-job-stage'),account_manager_id:c.account_manager_id,operator_id:value('gs-job-operator')||null,section:c.section,client_name:c.name,...(s.domain?{domain:s.domain}:{}),email:value('gs-job-email'),device_id:value('gs-job-device')||null,notes:value('gs-job-notes')});
      const job=r.state.setups.find(j=>j.id===id)||r.state.setups.at(-1);
      message(job.sync.status==='error'?'Domain saved; sheet sync needs attention. See its row.':'Domain saved.',job.sync.status==='error');
    });
    function options(){const owner=value('gs-job-operator');$('gs-job-device').innerHTML='<option value="">Allocate a device later</option>'+state().devices.filter(d=>d.id===s.device_id||available(d,id)&&d.member_id===owner).map(d=>`<option value="${esc(d.id)}" ${d.id===s.device_id?'selected':''}>${esc(label(d))} · ${esc(condition(d))} · ${esc(d.asset_tag)}</option>`).join('');if(s.purchase_device&&['Purchase made','Setup complete'].includes(s.stage))$('gs-job-device').disabled=true;}
    $('gs-job-operator').onchange=options;options();
    $('gs-job-device').required=false;
    $('gs-job-stage').onchange=()=>{$('gs-job-email').required=!s.domain||value('gs-job-stage')==='Setup complete';$('gs-job-operator').required=value('gs-job-stage')==='Setup complete';};$('gs-job-stage').onchange();
  }
  function addSetupDevice(id){
    const job=state().setups.find(s=>s.id===id);
    dialog('Add device',`<div class="gs-form-grid"><label>Setup owner<select id="gs-link-operator" required>${memberOptions(job.operator_id)}</select></label><label>Device<select id="gs-link-device" required></select></label></div>`,async()=>{
      await change('setup_device',{id,operator_id:value('gs-link-operator'),device_id:value('gs-link-device')});message('Device added.');
    });
    function options(){const owner=value('gs-link-operator');$('gs-link-device').innerHTML='<option value="">Select device</option>'+state().devices.filter(d=>available(d,id)&&d.member_id===owner).map(d=>`<option value="${esc(d.id)}">${esc(label(d))} · ${esc(condition(d))} · ${esc(d.asset_tag)}</option>`).join('');}
    $('gs-link-operator').onchange=options;options();
    if(job.sync?.status==='synced')$('gs-link-operator').disabled=true;
  }
  $('gs-capacity-form').onsubmit=ev=>{ev.preventDefault();const data={phone_setups:Number(value('gs-phone-capacity')),pc_setups:Number(value('gs-pc-capacity')),order_kind:value('gs-order-kind')};run(async()=>{await change('capacity_set',data);message('Setup limits saved. Device target and availability updated for the team.');});};
  $('gs-form').onsubmit=ev=>{ev.preventDefault();run(async()=>{await submitForm();$('gs-dialog').close();});};
  for(const id of ['gs-dialog-close','gs-dialog-cancel'])$(id).onclick=()=>$('gs-dialog').close();
  $('gs-refresh').onclick=()=>run(async()=>{await refresh();message('Team tracker refreshed.');});
  $('gs-add-setup').onclick=()=>{if(!selectedClient)return addClient();bulkOpen=true;renderSetups();$('gs-bulk-domains').focus();};
  $('gs-add-single').onclick=()=>editSetup();
  $('gs-edit-client').onclick=()=>addClient(selectedClient);
  $('gs-dmarc').onclick=()=>setupDmarc();
  $('gs-client-dns').onclick=()=>setupAllDomains();
  $('gs-delete-client').onclick=()=>{const c=state().clients.find(c=>c.id===selectedClient);if(c&&confirm(`Delete ${c.name} and its domains from Gmail Set? Google Sheet rows and device usage history will be retained.`))run(async()=>{await change('client_remove',{id:c.id});openClient(null);message('Client deleted.');});};
  $('gs-bulk-panel').onsubmit=ev=>{ev.preventDefault();run(async()=>{try{const r=await change('setup_bulk',{client_id:selectedClient,domains:value('gs-bulk-domains'),sender_name:value('gs-bulk-sender'),operator_id:value('gs-bulk-operator')||null});$('gs-bulk-domains').value='';bulkOpen=false;renderSetups();message(`${r.import_summary.added} domains added${r.import_summary.skipped?' · '+r.import_summary.skipped+' duplicates skipped':''}.`);$('gs-bulk-error').textContent='';}catch(err){$('gs-bulk-error').textContent=err.message;throw err;}});};
  $('gs-back-clients').onclick=()=>openClient(null);
  $('gs-client-list').onclick=e=>{const b=e.target.closest('[data-open-client]');if(b)openClient(b.dataset.openClient);};
  window.addEventListener('popstate',()=>{selectedClient=new URLSearchParams(location.hash.slice(1)).get('client');if(board)renderSetups();});$('gs-add-member').onclick=()=>editMember();$('gs-add-devices').onclick=()=>addDevices();
  $('gs-tab-setups').onclick=()=>tab('setups');$('gs-tab-devices').onclick=()=>tab('devices');
  for(const id of ['gs-search','gs-filter-stage','gs-filter-manager','gs-filter-section'])$(id).addEventListener('input',renderSetups);
  $('gs-stock').onclick=e=>{if(e.target.closest('[data-stock-add]')){tab('devices');addDevices();}};
  $('gs-team-list').onclick=e=>{
    const b=e.target.closest('button');if(!b)return;
    if(b.dataset.editMember)return editMember(b.dataset.editMember);
    if(b.dataset.addFor!==undefined)return addDevices(b.dataset.addFor);
    if(b.dataset.device)return editDevice(b.dataset.device);
    if(b.dataset.removeMember&&confirm('Remove this member? Their devices will move to the company pool. Purchase history stays saved.'))run(async()=>{await change('member_remove',{id:b.dataset.removeMember});message('Member removed. Devices are in the unallocated pool.');});
  };
  $('gs-setup-list').onclick=e=>{
    const b=e.target.closest('button');if(!b)return;
    if(b.dataset.addSetupDevice)return addSetupDevice(b.dataset.addSetupDevice);
    if(b.dataset.editSetup)return editSetup(b.dataset.editSetup);
    if(b.dataset.setupDns)return previewDnsSetup(b.dataset.setupDns);
    if(b.dataset.sync)return run(async()=>{message('Reconciling the sheet row…');const r=await api('/setups/'+b.dataset.sync+'/sync',{});accept(r);const job=r.state.setups.find(s=>s.id===b.dataset.sync);message(job.sync.status==='synced'?'Sending email verified Live in the strategy sheet.':job.sync.message,job.sync.status!=='synced');});
    if(b.dataset.removeSetup&&confirm('Delete this domain from Gmail Set? Google Sheet rows and device usage history will be retained.'))run(async()=>{await change('setup_remove',{id:b.dataset.removeSetup});message('Domain deleted.');});
  };
  $('gs-setup-list').onchange=e=>{
    const input=e.target,id=input.dataset.setOperator||input.dataset.setDevice||input.dataset.setEmail||input.dataset.setStage;if(!id)return;
    const job=state().setups.find(s=>s.id===id);
    if(input.dataset.setStage){
      const next=input.value;if(next===job.stage)return;
      if(next==='Setup complete'&&(!job.email||!job.operator_id)){message(!job.email?'Add the sending email on this domain before completing setup.':'Select the setup person on this domain before completing setup.',true);renderSetups();const field=$('gs-setup-list').querySelector(`[data-set-${!job.email?'email':'operator'}="${id}"]`);field?.focus();return;}
      return run(async()=>{try{message(next==='Setup complete'?'Saving completion and syncing the Live sheet row…':'Saving setup stage…');const r=await change('setup_stage',{id,stage:next});const updated=r.state.setups.find(s=>s.id===id);message(next==='Setup complete'&&updated.sync.status==='error'?'Setup complete and saved. Sheet sync needs attention; see its row.':next==='Setup complete'?`Setup complete · sheet order ${updated.order}.`:'Stage updated.',next==='Setup complete'&&updated.sync.status==='error');}finally{renderSetups();}});
    }
    if(input.dataset.setEmail&&!input.checkValidity()){input.reportValidity();return;}
    const patch=input.dataset.setOperator?{operator_id:input.value||null}:input.dataset.setDevice?{device_id:input.value||null}:{email:input.value.trim()};
    run(async()=>{try{if(job.sync?.status==='synced'&&input.dataset.setDevice)await change('setup_device',{id,operator_id:job.operator_id,device_id:input.value});else await change('setup_assign',{id,...patch});message('Domain updated.');}finally{renderSetups();}});
  };
  $('gs-check-sheet').onclick=()=>run(async()=>{const r=await api('/check-sheet',{});$('gs-sheet-status').textContent=r.message;message(r.ready?'Sheet write access is ready.':'Sheet write access needs attention.',!r.ready);});
  // Refresh shared counts without disturbing an open edit form or active request.
  setInterval(()=>{if(!busy&&!$('gs-dialog').open&&!document.hidden&&!editing())refresh().catch(err=>message(err.message,true));},30000);
  document.addEventListener('visibilitychange',()=>{if(!document.hidden&&!busy&&!$('gs-dialog').open&&!editing())refresh().catch(err=>message(err.message,true));});
  run(refresh);
})();
