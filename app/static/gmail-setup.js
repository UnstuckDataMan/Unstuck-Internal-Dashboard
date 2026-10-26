/* Shared state is persisted on the server; the browser submits narrow actions. */
(() => {
  'use strict';
  const $=id=>document.getElementById(id),app=$('gs-app');if(!app)return;
  const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  let board=null,busy=false,submitForm=null,bulkOpen=false,bulkClient=null;
  let lockedControls=null;
  let activityVersion=0;
  let nextVersion=null,sessionCheck=null,reloading=false;
  const noteDrafts=new Map();
  const passwordDrafts=new Map();
  let sheetRefresh=null;
  let processingBatch=null,batchCurrent='',stopBatch=false;
  let selectedClient=new URLSearchParams(location.hash.slice(1)).get('client');
  const stages=['Not started','Started','Purchase made','Setup complete','Failed at QR','Failed at payment','Burned','Gmail deleted'];
  const inactive=stage=>['Burned','Gmail deleted'].includes(stage);
  const closed=stage=>stage==='Setup complete'||inactive(stage);
  const stageOptions=current=>stages.map(stage=>`<option ${stage===current?'selected':''} ${closed(current)&&![current,'Not started','Burned','Gmail deleted'].includes(stage)?'disabled':''}>${stage}</option>`).join('');
  const state=()=>board?.state||{members:[],devices:[],clients:[],setups:[]};
  const member=id=>state().members.find(m=>m.id===id),device=id=>state().devices.find(d=>d.id===id);
  const name=id=>member(id)?.name||'Unallocated';
  const progress=c=>board?.client_progress?.[c?.id]||{target:c?.domain_target||c?.planned_domains||0,complete:0,actual_todo:0,remaining:c?.domain_target||c?.planned_domains||0,replacements_needed:c?.domain_target||c?.planned_domains||0};
  const missingLabel=c=>state().setups.some(j=>j.client_id===c.id&&(inactive(j.stage)||j.archived))?'need replacing':'to add';
  const mastersFor=clientId=>state().setups.filter(j=>!j.archived&&j.client_id===clientId&&j.master_domain_id===j.id).sort((a,b)=>a.master_color-b.master_color);
  const masterGraph=clientId=>JSON.stringify(state().setups.filter(j=>!j.archived&&j.client_id===clientId).map(j=>[j.id,j.master_domain_id??null,j.master_color??null]).sort((a,b)=>a[0].localeCompare(b[0])));
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
  const sheetEligible=job=>['Setup complete','Gmail deleted'].includes(job.stage);
  const sheetPending=job=>Boolean(job.sync?.generation)&&['pending','running'].includes(job.sync?.status);
  const sheetSaved=job=>sheetEligible(job)&&job.sync?.status==='synced'&&(job.stage==='Gmail deleted'?job.sync.live===false:job.sync.live!==false);
  function updateClientButtons(){const c=state().clients.find(c=>c.id===selectedClient&&!c.archived),jobs=state().setups.filter(j=>j.client_id===c?.id&&!j.archived&&!inactive(j.stage));$('gs-client-dns').hidden=!c||app.dataset.dnsEnabled!=='true';$('gs-client-dns').disabled=busy||!jobs.length;$('gs-complete-client').hidden=!c;$('gs-complete-client').disabled=busy||!jobs.some(j=>j.stage!=='Setup complete');$('gs-client-sheets').hidden=!c;$('gs-client-sheets').disabled=busy||!state().setups.some(j=>j.client_id===c?.id&&!j.archived&&sheetEligible(j));}
  function message(text,error=false){$('gs-message').textContent=text;$('gs-message').classList.toggle('error',error);}
  async function api(path='',data,recoverSession=true){
    const r=await fetch('/api/gmail-setup'+path,{method:data===undefined?'GET':'POST',credentials:'same-origin',cache:'no-store',headers:{'Content-Type':'application/json','X-Domain-CSRF':app.dataset.csrf},...(data===undefined?{}:{body:JSON.stringify(data)})});
    let result;try{result=await r.json();}catch{throw Error('Unexpected server response. Refresh before retrying.');}
    if(!r.ok){
      // Only this header means the server rejected the request before any
      // change. Never replay an authorization failure or an uncertain write.
      if(data!==undefined&&recoverSession&&r.status===403&&r.headers.get('X-Gmail-Session')==='refresh'){
        await checkSessionVersion();return api(path,data,false);
      }
      const error=Error(typeof result.detail==='string'?result.detail:r.status===401?'Sign in again to continue.':'Check the form fields and retry.');error.status=r.status;error.revisionConflict=r.headers.get('X-Gmail-Conflict')==='revision';throw error;
    }
    return result;
  }
  function unsavedChanges(){
    if(noteDrafts.size||passwordDrafts.size||bulkOpen&&(value('gs-bulk-domains')||value('gs-bulk-sender')))return true;
    if(!board)return false;
    if(Number(value('gs-phone-capacity'))!==board.summary.phone_setups||Number(value('gs-pc-capacity'))!==board.summary.pc_setups||value('gs-order-kind')!==board.summary.order_kind)return true;
    return [...app.querySelectorAll('[data-set-email]')].some(el=>el.value.trim()!==(state().setups.find(j=>j.id===el.dataset.setEmail)?.email||''));
  }
  function applyUpdate(){
    if(!nextVersion)return;
    const safe=Boolean(board)&&!busy&&!editing()&&!$('gs-dialog').open&&!$('gs-renew-dialog').open&&!unsavedChanges();
    $('gs-update').hidden=false;$('gs-apply-update').disabled=!safe;
    if(safe&&!document.hidden&&!reloading){reloading=true;location.reload();}
  }
  function checkSessionVersion(){
    if(!sessionCheck)sessionCheck=(async()=>{
      const session=await api('/session',undefined,false);
      app.dataset.csrf=session.csrf;
      nextVersion=session.app_version!==app.dataset.appVersion?session.app_version:null;
      $('gs-update').hidden=!nextVersion;
      applyUpdate();return session;
    })().finally(()=>{sessionCheck=null;});
    return sessionCheck;
  }
  function accept(result){
    if(board&&result.revision<board.revision)return;board=result;
    for(const id of noteDrafts.keys()){
      const job=state().setups.find(j=>j.id===id);
      if(!job||job.archived||state().clients.find(c=>c.id===job.client_id)?.archived)noteDrafts.delete(id);
    }
    for(const id of passwordDrafts.keys())if(!state().clients.some(c=>c.id===id&&!c.archived))passwordDrafts.delete(id);
    render();
  }
  async function refresh({background=false}={}){
    const version=activityVersion,result=await api();
    // A read can finish after someone starts choosing a stage or saving it.
    if(background&&(version!==activityVersion||busy||editing()||$('gs-dialog').open||$('gs-renew-dialog').open||unsavedChanges()))return;
    accept(result);
  }
  function lockControls(){
    if(!lockedControls)return;
    for(const el of lockedControls.keys())if(!el.isConnected)lockedControls.delete(el);
    for(const el of [...app.querySelectorAll('button, input, select, textarea'),...$('gs-dialog').querySelectorAll('button')]){
      if(el.matches('[data-pause-dns]'))continue;
      // Redraws can create controls while a save or DNS request is running.
      if(!lockedControls.has(el))lockedControls.set(el,el.disabled);
      el.disabled=true;
    }
  }
  async function run(fn){
    if(busy)return;busy=true;activityVersion++;app.setAttribute('aria-busy','true');
    lockedControls=new Map();lockControls();
    try{await fn();}catch(err){message(err.message,true);if($('gs-dialog').open)$('gs-form-error').textContent=err.message;}
    finally{busy=false;app.removeAttribute('aria-busy');lockedControls.forEach((disabled,el)=>{if(el.isConnected)el.disabled=disabled;});lockedControls=null;updateClientButtons();renderDnsBatch(state().clients.find(c=>c.id===selectedClient));applyUpdate();}
  }
  function trackerInputs(source){
    const {events,...snapshot}=source;
    snapshot.setups=(source.setups||[]).map(job=>{const {sync,...inputs}=job;return inputs;});
    const ordered=value=>Array.isArray(value)?value.map(ordered):value&&typeof value==='object'?Object.fromEntries(Object.keys(value).sort().map(key=>[key,ordered(value[key])])):value;
    return JSON.stringify(ordered(snapshot));
  }
  async function change(action,data){
    const clientPassword=action==='client_password';
    const originalState=state();
    const original=(clientPassword?state().clients:state().setups).find(j=>j.id===data.id);
    const originalMasterGraph=action==='setup_master'?masterGraph(original?.client_id):null;
    const watched=action==='setup_stage'?['stage','operator_id','device_id','email','domain','client_id','archived','purchase_at','purchase_device','device_usage_resets']:
      action==='setup_assign'?['stage','domain','client_id','archived','purchase_at','purchase_device',...('operator_id'in data||'device_id'in data?['operator_id','device_id']:[]),...('email'in data?['email']:[])]:
      action==='setup_notes'?['notes','domain','client_id','archived']:
      action==='setup_master'?['master_domain_id','master_color','domain','client_id','archived']:
      clientPassword?['password_updated_at','has_password','archived']:null;
    for(let attempt=0;attempt<3;attempt++){
      try{const result=await api('/change',{revision:board.revision,action,data});accept(result);return result;}
      catch(error){
        if(!error.revisionConflict)throw error;
        const latest=await api(),current=(clientPassword?latest.state.clients:latest.state.setups).find(j=>j.id===data.id);
        accept(latest);
        if(action==='setup_stage'&&current&&!current.archived&&current.stage===data.stage)return latest;
        const unchanged=watched?original&&current&&!current.archived&&watched.every(key=>JSON.stringify(original[key]??null)===JSON.stringify(current[key]??null)):trackerInputs(originalState)===trackerInputs(latest.state);
        if(!unchanged||action==='setup_master'&&masterGraph(current.client_id)!==originalMasterGraph){
          throw Error(clientPassword?'This client password was updated by another team member. Review it before saving again.':watched?'This domain was updated by another team member. The latest saved values are shown; review them before changing it again.':error.message);
        }
        // Retry the narrow change only if its inputs still match. Other
        // domains and unrelated notes can change without discarding a stage.
        // Other actions may rebase only a sheet status/audit update.
        if(attempt===2)throw Error('The team tracker is busy. Your change was not saved; please try again.');
      }
    }
  }
  function tab(which){for(const view of ['setups','devices']){$('gs-'+view).hidden=view!==which;$('gs-tab-'+view).setAttribute('aria-selected',String(view===which));}}
  function render(){
    const s=board.summary,st=state();
    $('gs-stats').innerHTML=[['not_started','Not started'],['started','Started'],['purchased','Purchase made'],['complete','Setup complete']].map(([key,title])=>`<div class="gs-stat"><strong>${s[key]}</strong><span>${title}</span></div>`).join('')+`<div class="gs-stat gs-reserve-stat"><span>Setup capacity / required</span><div class="gs-reserve-ratio"><strong>${s.available_setups}</strong><span>/</span><strong id="gs-reserve-target">${s.target}</strong></div><b class="gs-shortfall ${s.low_stock?'low':'ready'}">${s.low_stock?`Order ${s.shortfall} more ${s.order_kind==='Phone'?'phones':'laptops / PCs'}`:'Setup capacity covered'}</b><p class="gs-reserve-calculation">Required: ${s.todo} to do${s.planned_todo?` (${s.actual_todo} actual + ${s.planned_todo} to add / replace)`:""} + ${s.contingency} contingency = ${s.target} setups.<br>Available: ${s.phone_available} phone + ${s.pc_available} laptop / PC${s.purchased_setups?` + ${s.purchased_setups} already purchased`:""} = ${s.available_setups} setups.<br>${s.available_setups>s.target?`Surplus: ${s.available_setups} − ${s.target} = ${s.available_setups-s.target} setup slots.`:`Gap: ${s.target} − ${s.available_setups} = ${s.setup_shortfall} setups.`}<br>${s.setup_shortfall?`${s.setup_shortfall} ÷ ${s.order_capacity}, rounded up = ${s.shortfall} new ${s.order_kind==='Phone'?'phones':'laptops / PCs'} (${s.shortfall*s.order_capacity} extra setup slots).`:`No new devices needed. New ${s.order_kind==='Phone'?'phones':'laptops / PCs'} cover ${s.order_capacity} setups each.`}</p></div>`;
    $('gs-phone-capacity').value=s.phone_setups;$('gs-pc-capacity').value=s.pc_setups;$('gs-order-kind').value=s.order_kind;
    $('gs-device-rules').textContent=`Phones: ${s.phone_setups} setups each · Laptops / PCs: ${s.pc_setups} setups each.`;
    $('gs-stock').hidden=!s.low_stock;
    $('gs-stock').innerHTML=`<div><strong>Order ${s.shortfall} more ${s.order_kind==='Phone'?'phones':'laptops / PCs'}</strong><small>${s.available_setups} setup capacity / ${s.target} required · ${s.todo} to do${s.planned_todo?` including ${s.planned_todo} to add / replace`:""} + ${s.contingency} contingency setups. ${s.setup_shortfall} setups short; ${s.order_capacity} per new ${s.order_kind==='Phone'?'phone':'laptop / PC'}, rounded up.</small></div><button data-stock-add>Add inventory</button>`;
    const chosen=$('gs-filter-manager').value;$('gs-filter-manager').innerHTML='<option value="">Everyone</option>'+st.members.map(m=>`<option value="${esc(m.id)}">${esc(m.name)}${m.active?'':' (removed)'}</option>`).join('');$('gs-filter-manager').value=chosen;
    renderSetups();renderTeam();$('gs-updated').textContent=board.updated_at?`Updated ${new Date(board.updated_at).toLocaleString('en-GB')}`:'Add your actual device inventory to get started';lockControls();
  }
  function openClient(id){
    selectedClient=id||null;
    history.pushState(null,'',location.pathname+location.search+(id?'#client='+encodeURIComponent(id):''));
    renderSetups();
  }
  function sheetStatus(job){
    const sync=job.sync||{},eligible=sheetEligible(job),pending=sheetPending(job),saved=sheetSaved(job);
    if(!eligible&&!sync.row&&!pending&&!sync.message)return '';
    const safeLink=sync.url&&/^https:\/\/docs\.google\.com\/spreadsheets\/d\/[A-Za-z0-9_-]+\/edit#gid=\d+&range=A\d+$/.test(sync.url)?sync.url:null;
    const error=sync.status==='error',noRow=sync.status==='synced'&&sync.live===false&&(sync.missing||!sync.row);
    const title=pending?'Backing up to Sheets…':error?'Backup needs attention':noRow?'No sheet entry to update':saved?`Backed up · ${job.stage==='Gmail deleted'?'Not live':'Live'}`:eligible?'Not backed up yet':'Sheet entry retained';
    const details=[sync.tab,sync.row?`row ${sync.row}`:''].filter(Boolean).join(' · ');
    return `<div class="gs-sheet-status ${pending?'pending':error?'error':saved?'saved':''}" role="status"><strong>${esc(title)}</strong>${details?`<span>${esc(details)}</span>`:''}${safeLink?`<a href="${esc(safeLink)}" target="_blank" rel="noopener noreferrer">Open sheet row ↗</a>`:''}${(eligible||error&&sync.generation)&&!pending&&!saved?`<button class="gs-small" data-sync="${esc(job.id)}">${error?'Retry backup':'Back up to Sheets'}</button>`:''}${error&&sync.message?`<small>${esc(sync.message)}</small>`:''}</div>`;
  }
  function renderSheetSummary(client){
    const el=$('gs-sheet-backup-status'),jobs=state().setups.filter(j=>j.client_id===client?.id&&!j.archived&&sheetEligible(j));
    el.hidden=!client||!jobs.length;if(el.hidden){el.textContent='';return;}
    const saved=jobs.filter(sheetSaved).length,pending=jobs.filter(sheetPending).length,errors=jobs.filter(j=>j.sync?.status==='error').length,waiting=jobs.length-saved-pending-errors;
    el.textContent='Sheet backup: '+[saved?`${saved} backed up`:'',pending?`${pending} backing up…`:'',errors?`${errors} need attention`:'',waiting?`${waiting} not backed up`:''].filter(Boolean).join(' · ');
    el.classList.toggle('error',errors>0);
  }
  function renderDomain(s){
    const d=device(s.device_id),p=s.purchase_device,sync=s.sync||{},kind=s.section==='client'?'Client':'Biz Dev';
    const failure=s.stage.startsWith('Failed at'),published=Boolean(s.sheet_identity_locked??(sync.status==='synced')),burned=s.stage==='Burned',deleted=s.stage==='Gmail deleted',retired=inactive(s.stage);
    const cardStage=deleted?' gs-domain-deleted':burned||failure?' gs-domain-danger':s.stage==='Setup complete'?' gs-domain-complete':['Started','Purchase made'].includes(s.stage)?' gs-domain-started':'';
    const dns=s.dmarc&&s.dmarc.email===s.email&&s.dmarc.domain===s.domain?s.dmarc:null;
    const setup=s.dns_setup?.domain===s.domain&&s.dns_setup?.email===s.email?s.dns_setup:null;
    const setupReady=setup?.status==='verified'&&setup.configuration_version===2;
    const planning=['Not started','Started'].includes(s.stage);
    const ownerLocked=retired||!planning&&(published||s.stage==='Setup complete'||Boolean(p&&s.stage==='Purchase made')),deviceLocked=retired||!planning&&Boolean(d&&(published||p&&['Purchase made','Setup complete'].includes(s.stage)));
    const options='<option value="">Device not added yet</option>'+state().devices.filter(x=>x.id===s.device_id||available(x,s.id)&&x.member_id===s.operator_id).map(x=>`<option value="${esc(x.id)}" ${x.id===s.device_id?'selected':''}>${esc(label(x)+' · '+condition(x))}</option>`).join('');
    const masters=mastersFor(s.client_id),root=masters.find(m=>m.id===s.master_domain_id),isMaster=root?.id===s.id;
    const groupColour=root&&Number.isInteger(root.master_color)&&root.master_color>=0&&root.master_color<5?root.master_color:null;
    const crown=isMaster?'<svg class="gs-master-crown" role="img" aria-label="Master domain" viewBox="0 0 24 24"><path d="m3 6 5 4 4-7 4 7 5-4-2 13H5Z" fill="currentColor"/><path d="M6 22h12" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>':'';
    const masterOptions=`<option value="">No master domain</option><option value="${esc(s.id)}" ${isMaster?'selected':''} ${!isMaster&&masters.length>=5?'disabled':''}>${isMaster?'Master domain':masters.length>=5?'Master domain · 5 of 5 used':'Make master domain'}</option>`+masters.filter(m=>m.id!==s.id).map(m=>`<option value="${esc(m.id)}" ${root?.id===m.id?'selected':''}>Under ${esc(m.domain)}</option>`).join('');
    return `<article class="gs-domain${cardStage}${groupColour===null?'':' gs-master-'+groupColour}" data-setup="${esc(s.id)}"${groupColour===null?'':` data-master-root="${esc(root.id)}" data-master-color="${groupColour}"`}><div class="gs-domain-head"><div class="gs-domain-name"><h4>${esc(s.domain)}${crown}</h4>${s.order&&s.stage==='Setup complete'?`<small class="gs-order">Completion #${s.order}</small>`:''}</div><div class="gs-domain-tools"><label class="gs-master-choice">Master domain<select data-set-master="${esc(s.id)}" aria-label="Master domain for ${esc(s.domain)}">${masterOptions}</select></label><div class="gs-domain-buttons">${!published&&!retired?`<button class="gs-small" data-edit-setup="${esc(s.id)}">Details</button>`:''}<button class="gs-small gs-danger" data-remove-setup="${esc(s.id)}">Remove</button></div></div></div>
      <div class="gs-domain-controls"><label>Email<input type="email" data-set-email="${esc(s.id)}" value="${esc(s.email)}" placeholder="Add sending email" maxlength="254" ${retired||published||s.stage==='Setup complete'?'disabled':''}></label><label>Setup person<select data-set-operator="${esc(s.id)}" ${ownerLocked?'disabled':''}>${memberOptions(s.operator_id)}</select></label><label class="gs-device-choice">${d?`<strong class="gs-phone-number ${deviceClass(d)}">${esc((p?.id===d.id?p.label:d.kind+' '+d.number))}</strong>`:'Device'}<select data-set-device="${esc(s.id)}" aria-label="Device for ${esc(s.domain)}" ${deviceLocked?'disabled':''}>${options}</select></label><label class="gs-stage-control">Stage<select data-set-stage="${esc(s.id)}" aria-label="Stage for ${esc(s.domain)}">${stageOptions(s.stage,Boolean(s.purchase_at))}</select></label></div>
      ${retired||!d||failure||d?.status==='Blocked'||d&&exhausted(d)?`<div class="gs-domain-meta">${burned?'<span class="gs-badge burned">Burned</span>':deleted?'<span class="gs-badge deleted">Gmail deleted</span>':''}${!d?'<span class="gs-badge purchased">Device not added yet</span>':''}${failure?`<span class="gs-badge blocked">${esc(s.stage)}</span>`:''}${d?.status==='Blocked'?`<span class="gs-badge blocked">${esc(condition(d))} · do not use</span>`:d&&exhausted(d)?`<span class="gs-badge exhausted">${esc(condition(d))}</span>`:''}</div>`:''}
      ${!retired&&(app.dataset.dnsEnabled==='true'||dns||setup)?`<div class="gs-domain-dmarc">${dns?`<span class="gs-badge ${dns.status==='verified'?'complete':'blocked'}" title="${esc(dns.message)}">DMARC ${dns.status==='verified'?'verified':'needs attention'}</span>`:''}${app.dataset.dnsEnabled==='true'?`<button class="gs-small" data-setup-dns="${esc(s.id)}">${setupReady?'Recheck setup':'One-click setup'}</button>`:''}${setup?`<span class="gs-badge ${setupReady?'complete':'blocked'}" title="${esc(setup.message)}">${setupReady?'DNS ready':setup.status==='verified'?'Google mail setup needed':'DNS setup needs attention'}</span>`:''}</div>`:''}
      <div class="gs-domain-notes"><label>Notes<textarea rows="2" maxlength="1000" data-set-notes="${esc(s.id)}" aria-label="Notes for ${esc(s.domain)}" placeholder="Add a note…">${esc(noteDrafts.has(s.id)?noteDrafts.get(s.id):s.notes)}</textarea></label><button class="gs-small" data-save-notes="${esc(s.id)}">Save notes</button><small data-notes-status="${esc(s.id)}" role="status">${noteDrafts.has(s.id)?'Unsaved changes':''}</small></div>
      <details class="gs-domain-details"><summary>Account details</summary><dl class="gs-facts"><div><dt>Account manager</dt><dd>${esc(name(s.account_manager_id))}</dd></div><div><dt>Purchase asset</dt><dd>${esc(p?.asset_tag||d?.asset_tag||'Not allocated')}</dd></div><div><dt>Sheet destination</dt><dd>${kind}</dd></div></dl>${(s.device_failures||[]).length?`<p class="gs-help">${s.device_failures.map(f=>esc(`${f.stage} · ${f.device.label?f.device.label+' · '+f.device.member+' · '+f.device.asset_tag:'Device not recorded'}`)).join('<br>')}</p>`:''}</details><div class="gs-sync" data-sheet-status="${esc(s.id)}">${sheetStatus(s)}</div></article>`;
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
    $('gs-client-caption').textContent=current?`${name(current.account_manager_id)} · ${current.section==='client'?'Client':'Biz Dev'}`:'Choose a client to view its domains and setup stages.';
    renderClientInfo(current);
    $('gs-add-setup').textContent=current?'+ Add domains':'+ Client';
    $('gs-edit-client').hidden=!current;$('gs-delete-client').hidden=!current;
    updateClientButtons();renderDnsBatch(current);renderSheetSummary(current);
    if(current&&bulkClient!==current.id){bulkClient=current.id;bulkOpen=!all.some(s=>s.client_id===current.id);$('gs-bulk-domains').value='';$('gs-bulk-sender').value='';$('gs-bulk-operator').innerHTML=memberOptions(member(current.operator_id)?.active?current.operator_id:'');}
    $('gs-bulk-panel').hidden=!current||!bulkOpen;
    const visible=clients.slice().sort((a,b)=>a.name.localeCompare(b.name)).filter(c=>{
      const domains=all.filter(s=>s.client_id===c.id);
      return domains.some(matches)||(!domains.length&&!stage||progress(c).replacements_needed>0&&(!stage||stage==='Not started'))&&(!manager||c.account_manager_id===manager)&&(!section||c.section===section)&&(!search||c.name.toLowerCase().includes(search));
    });
    const clientToDo=c=>progress(c).remaining;
    const activeClients=visible.filter(c=>clientToDo(c)>0),completedClients=visible.filter(c=>clientToDo(c)===0);
    const clientCard=c=>{
      const p=progress(c),burned=all.filter(j=>j.client_id===c.id&&j.stage==='Burned').length,deleted=all.filter(j=>j.client_id===c.id&&j.stage==='Gmail deleted').length;
      return `<button class="gs-client-card" data-open-client="${esc(c.id)}"><div><strong>${esc(c.name)}</strong><small>${esc(name(c.account_manager_id))} · ${c.section==='client'?'Client':'Biz Dev'}</small></div><div class="gs-client-counts"><span>${p.target} target</span><span class="gs-client-remaining ${p.remaining?'':'is-covered'}"><b>${p.remaining}</b> to complete</span><span class="gs-complete-count">${p.complete} complete</span>${p.replacements_needed?`<span class="gs-replacement-count" data-client-replacements="${esc(c.id)}"><b>${p.replacements_needed}</b> ${missingLabel(c)}</span>`:''}${burned?`<span class="gs-burned-count">${burned} burned</span>`:''}${deleted?`<span class="gs-deleted-count">${deleted} deleted</span>`:''}<span aria-hidden="true">→</span></div></button>`;
    };
    $('gs-client-list').innerHTML=visible.length?`${activeClients.length?`<div class="gs-client-group" data-client-group="active">${activeClients.map(clientCard).join('')}</div>`:''}${completedClients.length?`<section class="gs-client-group" data-client-group="complete" aria-label="Completed clients"><h4 class="gs-list-heading gs-complete-heading">Completed <span>${completedClients.length}</span></h4>${completedClients.map(clientCard).join('')}</section>`:''}`:'<div class="gs-empty"><strong>No clients to show</strong>Add a client first, then add its domains.</div>';
    if(current){
      const domains=all.filter(s=>s.client_id===current.id&&matches(s));
      const pending=domains.filter(s=>!closed(s.stage)),complete=domains.filter(s=>s.stage==='Setup complete'),burned=domains.filter(s=>s.stage==='Burned'),deleted=domains.filter(s=>s.stage==='Gmail deleted');
      const order=(a,b)=>(a.order||0)-(b.order||0)||a.domain.localeCompare(b.domain);
      const pendingOrder=(a,b)=>(a.created_at||'').localeCompare(b.created_at||'');
      $('gs-setup-list').innerHTML=domains.length?`${pending.length?`<h4 class="gs-list-heading">In progress <span>${pending.length}</span></h4>${pending.sort(pendingOrder).map(renderDomain).join('')}`:''}${complete.length?`<h4 class="gs-list-heading gs-complete-heading">Live domains (complete) <span>${complete.length}</span></h4>${complete.sort(order).map(renderDomain).join('')}`:''}${burned.length?`<h4 class="gs-list-heading gs-burned-heading">Burned <span>${burned.length}</span></h4>${burned.sort((a,b)=>(a.burned_at||'').localeCompare(b.burned_at||'')).map(renderDomain).join('')}`:''}${deleted.length?`<h4 class="gs-list-heading gs-deleted-heading">Deleted domains <span>${deleted.length}</span></h4>${deleted.sort((a,b)=>(a.gmail_deleted_at||'').localeCompare(b.gmail_deleted_at||'')).map(renderDomain).join('')}`:''}`:`<div class="gs-empty"><strong>No domains to show</strong>${progress(current).remaining?`${progress(current).remaining} domains still needed. Add domains to continue.`:'Add a domain to this client or change your filters.'}</div>`;
    }else $('gs-setup-list').innerHTML='';
    lockControls();
  }
  function renderClientInfo(client){
    const el=$('gs-client-info');el.hidden=!client;if(!client){el.innerHTML='';return;}
    const p=progress(client),url=/^https?:\/\//i.test(client.website_url||'')?client.website_url:null;
    const password=`<dt class="gs-password-label">Password</dt><dd><div class="gs-client-password"><input id="gs-client-password" type="password" autocomplete="new-password" maxlength="1024" aria-label="Client password" placeholder="${client.has_password?'Password saved':'Add password'}" value="${esc(passwordDrafts.get(client.id)||'')}"><div class="gs-password-actions"><button id="gs-save-client-password" class="gs-small" ${passwordDrafts.has(client.id)?'':'disabled'}>Save</button><button id="gs-show-client-password" class="gs-small" aria-pressed="false" ${client.has_password||passwordDrafts.has(client.id)?'':'disabled'}>Show</button>${client.has_password?'<button id="gs-remove-client-password" class="gs-small gs-danger">Remove</button>':''}</div><small id="gs-password-status" role="status">${passwordDrafts.has(client.id)?'Unsaved changes':client.has_password?'Saved':''}</small></div></dd>`;
    const card=`<dt class="gs-card-last4-label">Card last 4 digits</dt><dd><span class="gs-card-last4${client.card_last4?'':' is-empty'}" data-client-card-last4="${esc(client.id)}">${esc(client.card_last4||'Not added')}</span></dd>`;
    el.innerHTML=`<div class="gs-client-progress"><span><strong>${p.target}</strong> target</span><span><strong>${p.complete}</strong> complete</span><span class="gs-client-remaining ${p.remaining?'':'is-covered'}"><strong>${p.remaining}</strong> still needed</span>${p.replacements_needed?`<span class="gs-replacement-count" data-client-replacements="${esc(client.id)}"><b>${p.replacements_needed}</b> ${missingLabel(client)}</span>`:''}</div><dl class="gs-client-contacts"><div><dt>Sender’s full name</dt><dd>${esc(client.sender_full_name||'Not added')}</dd>${password}</div><div><dt>Forwarding website</dt><dd>${url?`<a href="${esc(url)}" target="_blank" rel="noopener noreferrer">${esc(url)}</a>`:'Not added'}</dd></div><div><dt>Account manager email</dt><dd>${esc(client.account_manager_email||'Not added')}</dd>${card}</div></dl>`;
  }
  function completeClient(){
    const client=state().clients.find(c=>c.id===selectedClient);if(!client)return;
    const items=state().setups.filter(j=>j.client_id===client.id&&!j.archived&&!closed(j.stage));if(!items.length)return;
    dialog('Mark all as complete',`<p>Mark all <strong>${items.length} unfinished ${items.length===1?'domain':'domains'}</strong> for <strong>${esc(client.name)}</strong> as complete?</p><p>This moves them to Live domains (complete) and backs them up to the Master Strategy sheet as Live. Burned and Gmail deleted domains stay in their sections. Only continue when their Gmail setups are finished.</p>`,async()=>{
      await change('client_complete',{id:client.id});
      $('gs-filter-stage').value='';renderSetups();$('gs-dialog').close();
      message(`${items.length} domains marked complete. Sheet backup queued.`);
    });$('gs-dialog-save').textContent='Mark all as complete';
  }
  function backupClient(){
    const client=state().clients.find(c=>c.id===selectedClient&&!c.archived);if(!client)return;
    return run(async()=>{
      for(let attempt=0;attempt<3;attempt++){
        try{accept(await api('/clients/'+client.id+'/sync',{revision:board.revision}));break;}
        catch(error){
          if(!error.revisionConflict||attempt===2)throw error;
          // A finished backup can advance the board revision between clicks.
          // Requeue this client's current accounts without resubmitting edits.
          await refresh();
          if(!state().clients.some(c=>c.id===client.id&&!c.archived))throw Error('This client was removed.');
        }
      }
      message('Sheet backup queued for this client’s completed and deleted accounts. Progress is shown on each domain.');
    });
  }
  function removeSetup(id){
    const job=state().setups.find(j=>j.id===id&&!j.archived);if(!job)return;
    dialog('Remove domain from tracker?',`<div class="gs-remove-warning"><strong>If the Gmail subscription has been deleted, select “Gmail deleted” in the Stage dropdown instead.</strong><p>Cancel this popup and use that option to keep the domain in Deleted domains.</p></div><p class="gs-remove-explanation">Use Remove only to take <strong>${esc(job.domain)}</strong> out of the active tracker. It does not cancel Gmail or stop domain auto-renew. Google Sheet rows and device history are kept.</p>`,async()=>{
      await change('setup_remove',{id:job.id});message('Domain removed from the tracker.');
    });
    $('gs-dialog-save').textContent='Remove from tracker';$('gs-dialog-cancel').focus();
  }
  function renderDnsBatch(client){
    const batch=client?.dns_setup_batch,el=$('gs-dns-batch');el.hidden=!batch;
    if(!batch){el.innerHTML='';return;}
    const results=batch.results||{},ready=batch.items.filter(i=>results[i.id]?.status==='verified').length,attention=batch.items.filter(i=>results[i.id]?.status==='needs review').length,done=ready+attention,total=batch.items.length;
    const active=processingBatch===client.id,pending=done<total;
    const status=active?(stopBatch?'Pausing after the current domain…':`Setting up ${batchCurrent||'next domain'}…`):pending?(batch.paused_reason||'Paused. Resume to continue the waiting domains.'):'Finished';
    el.innerHTML=`<section class="gs-card gs-batch-card"><div class="gs-batch-header"><h4>All-domain setup</h4><strong>${done} / ${total} processed</strong></div><progress max="${total}" value="${done}" aria-label="Domain setup progress"></progress><p class="gs-batch-counts"><span>${ready} ready</span><span>${attention} need attention</span><span>${total-done} waiting</span></p><p class="gs-batch-status" role="status">${esc(status)}</p>${active?'<p class="gs-help">Keep this page open while setup runs. You can resume after a refresh.</p>':''}<div class="gs-actions">${app.dataset.dnsEnabled==='true'?(active?`<button data-pause-dns ${stopBatch?'disabled':''}>${stopBatch?'Pausing…':'Pause after this domain'}</button>`:pending?'<button data-resume-dns class="gs-primary">Resume setup</button>':''):''}</div><details><summary>Domain results</summary><ul class="gs-batch-results">${batch.items.map(i=>{const r=results[i.id]||{},ok=r.status==='verified',failed=r.status==='needs review';return `<li><strong>${esc(i.domain)}</strong><span class="gs-badge ${ok?'complete':failed?'blocked':'purchased'}">${ok?'Ready':failed?'Needs attention':active&&batchCurrent===i.domain||r.status==='running'?'In progress':'Waiting'}</span>${r.message?`<small>${esc(r.message)}</small>`:''}</li>`;}).join('')}</ul></details></section>`;
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
    const items=state().setups.filter(j=>j.client_id===client.id&&!j.archived&&!inactive(j.stage));if(!items.length)return;
    dialog('Set up all domains for '+client.name,`<p class="gs-dmarc-warning"><strong>Only proceed if you’ve completed purchase for all ${items.length} domains.</strong></p><label>Client website URL<input id="gs-batch-destination" inputmode="url" placeholder="https://clientwebsite.com" value="${esc(client.website_url||'')}" maxlength="2048" required></label><p>Set up all ${items.length} active domains for this client: remove default DNS, add Google MX/SPF, tracking CNAME and DMARC, then replace root forwarding with the URL above.</p><p>Domains run one at a time. Burned, Gmail deleted and removed domains are excluded. Keep this page open; saved progress can be resumed after a refresh.</p><details><summary>View ${items.length} domains</summary><ul class="gs-dmarc-preview">${items.map(i=>`<li><strong>${esc(i.domain)}</strong><small>${esc(i.email||'Sending email missing — needs attention')}</small></li>`).join('')}</ul></details>`,async()=>{
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
  function confirmAutoRenew(domain){
    const prompt=$('gs-renew-dialog');
    $('gs-renew-domain').textContent=domain||'this domain';prompt.returnValue='';
    return new Promise(resolve=>{prompt.addEventListener('close',()=>resolve(prompt.returnValue==='yes'),{once:true});prompt.showModal();});
  }
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
    const passwordField=id?'':field('Password (optional)','gs-client-new-password','','type="password" autocomplete="new-password" maxlength="1024"');
    dialog(id?'Edit client':'Add client',`<div class="gs-form-grid">${field('Client / company name','gs-client-name',c.name,'required maxlength="150"')}<label>Section<select id="gs-client-section" ${id?'disabled':''}><option value="client" ${c.section==='client'?'selected':''}>Client · 1) Emails - entry</option><option value="bizdev" ${c.section==='bizdev'?'selected':''}>Biz Dev · 1) BD Emails - entry</option></select></label>${field('Sender’s full name','gs-client-sender-full-name',c.sender_full_name,'maxlength="150"')}${field('Client URL for forwarding','gs-client-url',c.website_url,'inputmode="url" maxlength="2048" placeholder="https://clientwebsite.com"')}<label>Account manager<select id="gs-client-manager" required>${memberOptions(c.account_manager_id)}</select></label>${field('Account manager email','gs-client-manager-email',c.account_manager_email??member(c.account_manager_id)?.email??'','type="email" maxlength="254"')}<label>Default setup person (optional)<select id="gs-client-operator">${memberOptions(member(c.operator_id)?.active?c.operator_id:'')}</select></label><label>Domain target<input id="gs-client-target" type="number" inputmode="numeric" min="0" max="5000" step="1" value="${progress(c).target}"><small>Burned domains still need replacing to reach this target.</small></label>${passwordField}${field('Card last 4 digits (optional)','gs-client-card-last4',c.card_last4||'','type="text" inputmode="numeric" pattern="[0-9]{4}" maxlength="4" autocomplete="off"')}</div>`,async()=>{
      const existing=new Set(state().clients.map(c=>c.id));
      const password=id?'':$('gs-client-new-password').value;
      const r=await change('client_save',{...(id?{id}:{}),name:value('gs-client-name'),section:value('gs-client-section'),account_manager_id:value('gs-client-manager'),operator_id:value('gs-client-operator')||null,domain_target:Number(value('gs-client-target')||0),sender_full_name:value('gs-client-sender-full-name'),website_url:value('gs-client-url'),account_manager_email:value('gs-client-manager-email'),card_last4:$('gs-client-card-last4').value,...(password!==''?{password}:{})});
      if(!id)openClient(r.state.clients.find(c=>!existing.has(c.id)).id);message(id?'Client updated.':'Client added.');
    });
    $('gs-client-manager').onchange=()=>{$('gs-client-manager-email').value=member(value('gs-client-manager'))?.email||'';};
  }
  function editSetup(id){
    const s=state().setups.find(s=>s.id===id)||{};
    let selectedStage=s.stage||'Not started',autoRenewConfirmed=false;
    const c=state().clients.find(c=>c.id===(s.client_id||selectedClient));
    if(!c)return addClient();
    dialog(id?'Edit domain':`Add domain to ${c.name}`,`<div class="gs-form-grid"><label>Setup owner<select id="gs-job-operator">${memberOptions(s.operator_id||c.operator_id)}</select></label>${field('Email','gs-job-email',s.email,'type="email" maxlength="254"')}<label>Stage<select id="gs-job-stage">${stageOptions(s.stage||'Not started',Boolean(s.purchase_at))}</select></label><label class="gs-wide">Device (optional)<select id="gs-job-device"><option value="">Allocate a device later</option></select></label><label class="gs-wide">Notes (optional)<textarea id="gs-job-notes" maxlength="1000">${esc(noteDrafts.has(id)?noteDrafts.get(id):s.notes)}</textarea></label></div>`,async()=>{
      const r=await change('setup_save',{...(id?{id}:{}),client_id:c.id,stage:value('gs-job-stage'),...(autoRenewConfirmed?{confirmed_auto_renew_stopped:true}:{}),account_manager_id:c.account_manager_id,operator_id:value('gs-job-operator')||null,section:c.section,client_name:c.name,...(s.domain?{domain:s.domain}:{}),email:value('gs-job-email'),device_id:value('gs-job-device')||null,notes:value('gs-job-notes')});
      const job=r.state.setups.find(j=>j.id===id)||r.state.setups.at(-1);
      noteDrafts.delete(job.id);renderSetups();
      message(job.stage==='Setup complete'?'Setup complete. Sheet backup queued.':job.stage==='Gmail deleted'?'Moved to Deleted domains. The sheet entry will be kept with Live unticked.':'Domain saved.');
    });
    function options(){const owner=value('gs-job-operator');$('gs-job-device').innerHTML='<option value="">Allocate a device later</option>'+state().devices.filter(d=>d.id===s.device_id||available(d,id)&&d.member_id===owner).map(d=>`<option value="${esc(d.id)}" ${d.id===s.device_id?'selected':''}>${esc(label(d))} · ${esc(condition(d))} · ${esc(d.asset_tag)}</option>`).join('');if(s.purchase_device&&['Purchase made','Setup complete'].includes(s.stage))$('gs-job-device').disabled=true;}
    $('gs-job-operator').onchange=options;options();
    $('gs-job-device').required=false;
    function requiredFields(){$('gs-job-email').required=!s.domain||value('gs-job-stage')==='Setup complete';$('gs-job-operator').required=value('gs-job-stage')==='Setup complete';}
    $('gs-job-stage').onchange=async()=>{
      const next=value('gs-job-stage');
      if(next==='Gmail deleted'&&s.stage!=='Gmail deleted'){
        $('gs-job-stage').value=selectedStage;
        if(!await confirmAutoRenew(s.domain||value('gs-job-email').split('@').at(-1)))return;
        $('gs-job-stage').value=next;autoRenewConfirmed=true;
      }else autoRenewConfirmed=false;
      selectedStage=next;requiredFields();
    };requiredFields();
  }
  function addSetupDevice(id){
    const job=state().setups.find(s=>s.id===id);
    dialog('Add device',`<div class="gs-form-grid"><label>Setup owner<select id="gs-link-operator" required>${memberOptions(job.operator_id)}</select></label><label>Device<select id="gs-link-device" required></select></label></div>`,async()=>{
      await change('setup_device',{id,operator_id:value('gs-link-operator'),device_id:value('gs-link-device')});message('Device added.');
    });
    function options(){const owner=value('gs-link-operator');$('gs-link-device').innerHTML='<option value="">Select device</option>'+state().devices.filter(d=>available(d,id)&&d.member_id===owner).map(d=>`<option value="${esc(d.id)}">${esc(label(d))} · ${esc(condition(d))} · ${esc(d.asset_tag)}</option>`).join('');}
    $('gs-link-operator').onchange=options;options();
    if((job.sheet_identity_locked??(job.sync?.status==='synced'))&&!['Not started','Started'].includes(job.stage))$('gs-link-operator').disabled=true;
  }
  $('gs-capacity-form').onsubmit=ev=>{ev.preventDefault();const data={phone_setups:Number(value('gs-phone-capacity')),pc_setups:Number(value('gs-pc-capacity')),order_kind:value('gs-order-kind')};run(async()=>{await change('capacity_set',data);message('Setup limits saved. Device target and availability updated for the team.');});};
  $('gs-form').onsubmit=ev=>{ev.preventDefault();run(async()=>{await submitForm();$('gs-dialog').close();});};
  for(const id of ['gs-dialog-close','gs-dialog-cancel'])$(id).onclick=()=>$('gs-dialog').close();
  $('gs-refresh').onclick=()=>run(async()=>{await refresh();await checkSessionVersion();message('Team tracker refreshed.');});
  $('gs-apply-update').onclick=applyUpdate;
  $('gs-dialog').addEventListener('close',()=>setTimeout(applyUpdate,0));
  $('gs-renew-dialog').addEventListener('close',()=>setTimeout(applyUpdate,0));
  $('gs-add-setup').onclick=()=>{if(!selectedClient)return addClient();bulkOpen=true;renderSetups();$('gs-bulk-domains').focus();};
  $('gs-add-single').onclick=()=>editSetup();
  $('gs-edit-client').onclick=()=>addClient(selectedClient);
  $('gs-client-dns').onclick=()=>setupAllDomains();
  $('gs-complete-client').onclick=completeClient;
  $('gs-client-sheets').onclick=backupClient;
  $('gs-delete-client').onclick=()=>{const c=state().clients.find(c=>c.id===selectedClient);if(c&&confirm(`Delete ${c.name} and its domains from Gmail Setup Tool? Google Sheet rows and device usage history will be retained.`))run(async()=>{await change('client_remove',{id:c.id});openClient(null);message('Client deleted.');});};
  $('gs-bulk-panel').onsubmit=ev=>{ev.preventDefault();run(async()=>{try{const r=await change('setup_bulk',{client_id:selectedClient,domains:value('gs-bulk-domains'),sender_name:value('gs-bulk-sender'),operator_id:value('gs-bulk-operator')||null});$('gs-bulk-domains').value='';bulkOpen=false;renderSetups();message(`${r.import_summary.added} domains added${r.import_summary.skipped?' · '+r.import_summary.skipped+' duplicates skipped':''}.`);$('gs-bulk-error').textContent='';}catch(err){$('gs-bulk-error').textContent=err.message;throw err;}});};
  $('gs-back-clients').onclick=()=>openClient(null);
  $('gs-client-info').oninput=e=>{
    if(e.target.id!=='gs-client-password'||!selectedClient)return;
    if(e.target.value)passwordDrafts.set(selectedClient,e.target.value);else passwordDrafts.delete(selectedClient);
    $('gs-save-client-password').disabled=busy||!passwordDrafts.has(selectedClient);
    $('gs-show-client-password').disabled=busy||!e.target.value&&!state().clients.find(c=>c.id===selectedClient)?.has_password;
    $('gs-password-status').textContent=passwordDrafts.has(selectedClient)?'Unsaved changes':'';
  };
  $('gs-client-info').onclick=e=>{
    const button=e.target.closest('button'),client=state().clients.find(c=>c.id===selectedClient);if(!button||!client)return;
    const input=$('gs-client-password');
    if(button.id==='gs-save-client-password'){
      const password=input.value;if(!password)return;
      return run(async()=>{await change('client_password',{id:client.id,password});passwordDrafts.delete(client.id);renderClientInfo(state().clients.find(c=>c.id===selectedClient&&!c.archived));message('Client password saved.');});
    }
    if(button.id==='gs-remove-client-password')return run(async()=>{await change('client_password',{id:client.id,password:null});passwordDrafts.delete(client.id);renderClientInfo(state().clients.find(c=>c.id===selectedClient&&!c.archived));message('Client password removed.');});
    if(button.id==='gs-show-client-password'){
      if(input.type==='text'){
        input.type='password';button.textContent='Show';button.setAttribute('aria-pressed','false');if(!passwordDrafts.has(client.id))input.value='';return;
      }
      if(passwordDrafts.has(client.id)){input.type='text';button.textContent='Hide';button.setAttribute('aria-pressed','true');return;}
      return run(async()=>{const result=await api('/clients/'+client.id+'/password/reveal',{});input.value=result.password;input.type='text';button.textContent='Hide';button.setAttribute('aria-pressed','true');});
    }
  };
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
    if(b.dataset.sync)return run(async()=>{accept(await api('/setups/'+b.dataset.sync+'/sync',{}));message('Sheet backup queued.');});
    if(b.dataset.saveNotes)return run(async()=>{const id=b.dataset.saveNotes,notes=$('gs-setup-list').querySelector(`[data-set-notes="${id}"]`).value;await change('setup_notes',{id,notes});noteDrafts.delete(id);renderSetups();$('gs-setup-list').querySelector(`[data-notes-status="${id}"]`).textContent='Saved';message('Notes saved.');});
    if(b.dataset.removeSetup)return removeSetup(b.dataset.removeSetup);
  };
  $('gs-setup-list').oninput=e=>{const id=e.target.dataset.setNotes;if(!id)return;if(e.target.value===(state().setups.find(j=>j.id===id)?.notes||''))noteDrafts.delete(id);else noteDrafts.set(id,e.target.value);$('gs-setup-list').querySelector(`[data-notes-status="${id}"]`).textContent=noteDrafts.has(id)?'Unsaved changes':'';};
  $('gs-setup-list').onchange=async e=>{
    const input=e.target,id=input.dataset.setOperator||input.dataset.setDevice||input.dataset.setEmail||input.dataset.setStage||input.dataset.setMaster;if(!id)return;
    const job=state().setups.find(s=>s.id===id);
    if(input.dataset.setMaster){
      const next=input.value||null;if(next===(job.master_domain_id||null))return;
      const wasMaster=job.master_domain_id===job.id;
      return run(async()=>{try{await change('setup_master',{id,master_domain_id:next});message(wasMaster&&next!==id?'Master changed. Previously linked domains are now ungrouped.':next===id?'Master domain saved.':next?'Domain linked to its master.':'Domain ungrouped.');}finally{renderSetups();}});
    }
    if(input.dataset.setStage){
      const next=input.value;if(next===job.stage)return;
      if(next==='Gmail deleted'){input.value=job.stage;if(!await confirmAutoRenew(job.domain))return;}
      if(next==='Setup complete'&&(!job.email||!job.operator_id)){message(!job.email?'Add the sending email on this domain before completing setup.':'Select the setup person on this domain before completing setup.',true);renderSetups();const field=$('gs-setup-list').querySelector(`[data-set-${!job.email?'email':'operator'}="${id}"]`);field?.focus();return;}
      return run(async()=>{try{message('Saving setup stage…');await change('setup_stage',{id,stage:next,...(next==='Gmail deleted'?{confirmed_auto_renew_stopped:true}:{})});message(next==='Setup complete'?'Setup complete. Sheet backup queued.':next==='Gmail deleted'?'Moved to Deleted domains. The sheet entry will be kept with Live unticked.':'Stage updated.');}finally{renderSetups();}});
    }
    if(input.dataset.setEmail&&!input.checkValidity()){input.reportValidity();return;}
    const patch=input.dataset.setOperator?{operator_id:input.value||null}:input.dataset.setDevice?{device_id:input.value||null}:{email:input.value.trim()};
    run(async()=>{try{if((job.sheet_identity_locked??(job.sync?.status==='synced'))&&!['Not started','Started'].includes(job.stage)&&input.dataset.setDevice)await change('setup_device',{id,operator_id:job.operator_id,device_id:input.value});else await change('setup_assign',{id,...patch});message('Domain updated.');}finally{renderSetups();}});
  };
  // Open tabs check the running version too. Drafts and active writes delay
  // an upgrade; returning to an idle tab loads the same release for everyone.
  function poll(){
    if(document.hidden)return;
    checkSessionVersion().catch(err=>message(err.message,true));
    if(!busy&&!$('gs-dialog').open&&!$('gs-renew-dialog').open&&!editing()&&!unsavedChanges())refresh({background:true}).catch(err=>message(err.message,true));
  }
  setInterval(poll,30000);
  // Backups finish independently of stage saves. Reuse the guarded refresh so
  // a result cannot replace a dropdown selection, dialog or unsaved draft.
  setInterval(()=>{
    if(sheetRefresh||document.hidden||busy||$('gs-dialog').open||$('gs-renew-dialog').open||editing()||unsavedChanges()||!state().setups.some(j=>!j.archived&&sheetPending(j)))return;
    sheetRefresh=refresh({background:true}).catch(err=>message(err.message,true)).finally(()=>{sheetRefresh=null;});
  },2000);
  document.addEventListener('visibilitychange',poll);
  window.addEventListener('pageshow',event=>{if(event.persisted)poll();});
  run(async()=>{await refresh();await checkSessionVersion();});
})();
