/* Shared state is persisted on the server; the browser submits narrow actions. */
(() => {
  'use strict';
  const $=id=>document.getElementById(id),app=$('gs-app');if(!app)return;
  const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  let board=null,busy=false,submitForm=null;
  let selectedClient=new URLSearchParams(location.hash.slice(1)).get('client');
  const stages=['Not started','Purchase made','Setup complete','Failed at QR','Failed at payment'];
  const stageOptions=(current,purchased=false)=>stages.map(stage=>`<option ${stage===current?'selected':''} ${current==='Setup complete'&&stage!==current||stage==='Not started'&&(current==='Purchase made'||purchased)?'disabled':''}>${stage}</option>`).join('');
  const state=()=>board?.state||{members:[],devices:[],clients:[],setups:[]};
  const member=id=>state().members.find(m=>m.id===id),device=id=>state().devices.find(d=>d.id===id);
  const name=id=>member(id)?.name||'Unallocated';
  const memberOptions=(selected='',pool=false)=>(pool?'<option value="">Unallocated pool</option>':'<option value="">Select team member</option>')+state().members.filter(m=>m.active||m.id===selected).map(m=>`<option value="${esc(m.id)}" ${m.id===selected?'selected':''}>${esc(m.name)}${m.active?'':' (removed)'}</option>`).join('');
  const uses=d=>Math.max(d.use_count||0,d.used_at||d.status==='Used'?1:0);
  const exhausted=d=>d.kind==='Phone'&&uses(d)>=2;
  const available=(d,jobId)=>{const reserved=state().setups.filter(j=>!j.archived&&j.stage==='Not started'&&j.id!==jobId&&j.device_id===d.id).length;return !d.archived&&d.status!=='Blocked'&&(d.kind==='Phone'?uses(d)+reserved<2:reserved===0);};
  const deviceClass=d=>d.status==='Blocked'?'blocked':exhausted(d)?'exhausted phone':d.kind==='Phone'?(uses(d)?'used phone':'unused phone'):d.status.toLowerCase();
  const condition=d=>d.status==='Blocked'&&d.block_reason&&d.block_reason!=='Other'?'Block '+d.block_reason:d.status==='Blocked'?'Blocked':d.kind==='Phone'?`${uses(d)} / 2 setups${exhausted(d)?' · do not use':''}`:d.status;
  const conditions=[['Unused','Unused'],['Used','Used'],['Blocked|QR code','Block QR code'],['Blocked|Payment','Block Payment'],['Blocked|Other','Blocked (other)']];
  const conditionData=v=>({status:v.split('|')[0],block_reason:v.split('|')[1]||null});
  const label=d=>`${d.kind} ${d.number} · ${name(d.member_id)}`;
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
    const buttons=[...app.querySelectorAll('button, [data-set-stage]'),...$('gs-dialog').querySelectorAll('button')];const disabled=buttons.map(b=>b.disabled);buttons.forEach(b=>b.disabled=true);
    try{await fn();}catch(err){message(err.message,true);if($('gs-dialog').open)$('gs-form-error').textContent=err.message;}
    finally{busy=false;app.removeAttribute('aria-busy');buttons.forEach((b,i)=>{if(b.isConnected)b.disabled=disabled[i];});}
  }
  async function change(action,data){const result=await api('/change',{revision:board.revision,action,data});accept(result);return result;}
  function tab(which){for(const view of ['setups','devices']){$('gs-'+view).hidden=view!==which;$('gs-tab-'+view).setAttribute('aria-selected',String(view===which));}}
  function render(){
    const s=board.summary,st=state();
    $('gs-stats').innerHTML=[['not_started','Not started'],['purchased','Purchase made'],['complete','Setup complete'],['unused','Unused devices / 20']].map(([key,title])=>`<div class="gs-stat"><strong>${s[key]}</strong><span>${title}</span></div>`).join('');
    $('gs-stock').hidden=!s.low_stock;
    $('gs-stock').innerHTML=`<div><strong>Order new phones · ${s.unused} / ${s.target} unused devices in reserve</strong><small>${s.shortfall} more unused devices needed company-wide. Used and blocked devices are excluded.</small></div><button data-stock-add>Add inventory</button>`;
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
    const purchase=p&&p.id===s.device_id?`${p.label} · ${p.member}`:d?label(d):null;
    const safeLink=sync.url&&/^https:\/\/docs\.google\.com\/spreadsheets\/d\/[A-Za-z0-9_-]+\/edit#gid=\d+&range=A\d+$/.test(sync.url)?sync.url:null;
    return `<article class="gs-domain" data-setup="${esc(s.id)}"><div class="gs-domain-head"><div><h4>${esc(s.domain)}</h4><small>${esc(s.email||'Sending email not added yet')}</small></div><label class="gs-stage-control">Stage<select data-set-stage="${esc(s.id)}" aria-label="Stage for ${esc(s.domain)}">${stageOptions(s.stage,Boolean(s.purchase_at))}</select></label></div><div class="gs-domain-meta"><span>Setup: ${esc(name(s.operator_id))}</span>${purchase?`<span>Device: ${esc(purchase)}</span>`:'<span class="gs-badge purchased">Device not added yet</span>'}${s.stage.startsWith('Failed at')?`<span class="gs-badge blocked">${esc(s.stage)}</span>`:''}${d&&exhausted(d)&&d.status!=='Blocked'?'<span class="gs-badge exhausted">2 / 2 phone setups · do not use</span>':''}${d?.status==='Blocked'?`<span class="gs-badge blocked">${esc(condition(d))} · do not use</span>`:''}</div><div class="gs-actions">${!s.device_id&&!s.stage.startsWith('Failed at')?`<button class="gs-small" data-add-setup-device="${esc(s.id)}">Add device</button>`:''}${sync.status!=='synced'?`<button class="gs-small" data-edit-setup="${esc(s.id)}">Edit details</button>`:''}${s.stage==='Setup complete'&&sync.status!=='synced'?`<button class="gs-small" data-sync="${esc(s.id)}">Retry sheet sync</button>`:''}${sync.status!=='synced'?`<button class="gs-small gs-danger" data-remove-setup="${esc(s.id)}">Archive</button>`:''}</div><details class="gs-domain-details"><summary>Account details</summary><dl class="gs-facts"><div><dt>Account manager</dt><dd>${esc(name(s.account_manager_id))}</dd></div><div><dt>Purchase asset</dt><dd>${esc(p?.asset_tag||d?.asset_tag||'Not allocated')}</dd></div><div><dt>Sheet destination</dt><dd>${kind}</dd></div></dl>${(s.device_failures||[]).length?`<p class="gs-help">${s.device_failures.map(f=>esc(`${f.stage} · ${f.device.label?f.device.label+' · '+f.device.member+' · '+f.device.asset_tag:'Device not recorded'}`)).join('<br>')}</p>`:''}${s.notes?`<p class="gs-help">${esc(s.notes)}</p>`:''}</details><div class="gs-sync ${sync.status==='error'?'error':''}">${sync.status==='synced'?`Live in ${esc(sync.tab)} · row ${sync.row}${safeLink?` · <a href="${esc(safeLink)}" target="_blank" rel="noopener noreferrer">Open sheet row ↗</a>`:''}`:sync.status==='error'?`Sheet sync needs attention: ${esc(sync.message)}${safeLink?` · <a href="${esc(safeLink)}" target="_blank" rel="noopener noreferrer">Open affected sheet row ↗</a>`:''}`:s.stage==='Setup complete'?'Saved as complete · sheet sync pending':''}</div></article>`;
  }
  function renderSetups(){
    const search=$('gs-search').value.toLowerCase(),stage=$('gs-filter-stage').value,manager=$('gs-filter-manager').value,section=$('gs-filter-section').value;
    const all=state().setups.filter(s=>!s.archived);
    const matches=s=>(!stage||s.stage===stage)&&(!manager||s.account_manager_id===manager)&&(!section||s.section===section)&&(!search||[s.domain,s.email,s.client_name].join(' ').toLowerCase().includes(search));
    const clients=state().clients||[];
    const current=clients.find(c=>c.id===selectedClient);
    if(!current)selectedClient=null;
    $('gs-client-list').hidden=Boolean(current);$('gs-client-detail').hidden=!current;
    $('gs-back-clients').hidden=!current;
    $('gs-client-heading').textContent=current?current.name:'Clients';
    $('gs-client-caption').textContent=current?`${name(current.account_manager_id)} · ${current.section==='client'?'Client':'Biz Dev'} · Completed domains move to the bottom`:'Choose a client to view its domains and setup stages.';
    $('gs-add-setup').textContent=current?'+ Add domain':'+ Client';
    const visible=clients.slice().sort((a,b)=>a.name.localeCompare(b.name)).filter(c=>{
      const domains=all.filter(s=>s.client_id===c.id);
      return domains.length?domains.some(matches):!stage&&(!manager||c.account_manager_id===manager)&&(!section||c.section===section)&&(!search||c.name.toLowerCase().includes(search));
    });
    $('gs-client-list').innerHTML=visible.length?visible.map(c=>{
      const domains=all.filter(s=>s.client_id===c.id),complete=domains.filter(s=>s.stage==='Setup complete').length;
      return `<button class="gs-client-card" data-open-client="${esc(c.id)}"><div><strong>${esc(c.name)}</strong><small>${esc(name(c.account_manager_id))} · ${c.section==='client'?'Client':'Biz Dev'}</small></div><div class="gs-client-counts"><span>${domains.length} ${domains.length===1?'domain':'domains'}</span><span>${domains.length-complete} to do</span><span class="gs-complete-count">${complete} complete</span><span aria-hidden="true">→</span></div></button>`;
    }).join(''):'<div class="gs-empty"><strong>No clients to show</strong>Add a client first, then add its domains.</div>';
    if(current){
      const domains=all.filter(s=>s.client_id===current.id&&matches(s));
      const pending=domains.filter(s=>s.stage!=='Setup complete'),complete=domains.filter(s=>s.stage==='Setup complete');
      const order=(a,b)=>a.domain.localeCompare(b.domain);
      $('gs-setup-list').innerHTML=domains.length?`${pending.length?`<h4 class="gs-list-heading">In progress <span>${pending.length}</span></h4>${pending.sort(order).map(renderDomain).join('')}`:''}${complete.length?`<h4 class="gs-list-heading gs-complete-heading">Complete <span>${complete.length}</span></h4>${complete.sort(order).map(renderDomain).join('')}`:''}`:'<div class="gs-empty"><strong>No domains to show</strong>Add a domain to this client or change your filters.</div>';
    }else $('gs-setup-list').innerHTML='';
  }
  function renderTeam(){
    const active=state().members.filter(m=>m.active),pool={id:null,name:'Unallocated pool',email:''};
    $('gs-team-list').innerHTML=[...active,pool].map(m=>{
      const ds=state().devices.filter(d=>!d.archived&&d.member_id===m.id).sort((a,b)=>a.kind.localeCompare(b.kind)||a.number-b.number);
      return `<article class="gs-card" data-member="${esc(m.id||'pool')}"><div class="gs-card-head"><div><h4>${esc(m.name)}</h4><small>${esc(m.email)}</small></div><div class="gs-actions">${m.id?`<button class="gs-small" data-edit-member="${esc(m.id)}">Edit</button><button class="gs-small gs-danger" data-remove-member="${esc(m.id)}">Remove</button>`:''}<button class="gs-small" data-add-for="${esc(m.id||'')}">+ Devices</button></div></div><div class="gs-device-counts"><span><strong>${ds.filter(d=>d.kind==='Phone').length}</strong> phones${m.id?' / 10':''}</span><span><strong>${ds.filter(d=>d.kind==='PC').length}</strong> PCs${m.id?' / 10':''}</span><span><strong>${ds.filter(d=>d.status==='Unused'&&!uses(d)).length}</strong> unused</span><span><strong>${ds.filter(d=>d.status==='Blocked'||exhausted(d)).length}</strong> unavailable</span></div>${ds.length?`<div class="gs-device-grid">${ds.map(d=>`<button class="gs-device ${deviceClass(d)}" data-device="${esc(d.id)}" aria-label="${esc(label(d)+' · '+condition(d)+' · '+d.asset_tag)}">${esc(d.kind+' '+d.number)} · ${esc(condition(d))}<small>${esc(d.asset_tag)}</small></button>`).join('')}</div>`:'<p class="gs-help">No devices allocated.</p>'}</article>`;
    }).join('');
  }
  function dialog(title,html,onSubmit){$('gs-dialog-title').textContent=title;$('gs-form-fields').innerHTML=html;$('gs-form-error').textContent='';submitForm=onSubmit;$('gs-dialog').showModal();}
  const value=id=>$(id).value.trim();
  const field=(title,id,val='',extra='')=>`<label>${title}<input id="${id}" value="${esc(val)}" ${extra}></label>`;
  function editMember(id){const m=member(id)||{};dialog(id?'Edit team member':'Add team member',`<div class="gs-form-grid">${field('Name','gs-member-name',m.name,'required maxlength="100"')}${field('Team email (optional)','gs-member-email',m.email,'type="email" maxlength="254"')}</div>`,async()=>{await change('member_save',{...(id?{id}:{}),name:value('gs-member-name'),email:value('gs-member-email')});message('Team member saved.');});}
  function addDevices(owner=''){
    const counts=Array.from({length:11},(_,i)=>`<option value="${i}">${i}</option>`).join('');
    dialog('Add device inventory',`<div class="gs-form-grid"><label class="gs-wide">Allocate to<select id="gs-device-owner">${memberOptions(owner,true)}</select></label><label>Number of phones<select id="gs-device-phones">${counts}</select></label><label>Number of PCs<select id="gs-device-pcs">${counts}</select></label><label class="gs-wide">Current condition<select id="gs-device-condition">${conditions.map(([v,l])=>`<option value="${esc(v)}">${l}</option>`).join('')}</select></label></div>`,async()=>{await change('devices_add',{member_id:value('gs-device-owner')||null,phones:Number(value('gs-device-phones')),pcs:Number(value('gs-device-pcs')),...conditionData(value('gs-device-condition'))});message('Device inventory added.');});
  }
  function editDevice(id){const d=device(id);dialog(label(d),`<div class="gs-form-grid">${field('Physical asset label','gs-asset-tag',d.asset_tag,'required maxlength="80"')}<label>Allocate to<select id="gs-transfer-owner">${memberOptions(d.member_id,true)}</select></label><label>Condition<select id="gs-asset-status">${conditions.map(([v,l])=>`<option value="${esc(v)}" ${v===(d.status==='Blocked'?'Blocked|'+(d.block_reason||'Other'):d.status)?'selected':''} ${d.status==='Blocked'&&!v.startsWith('Blocked')||d.used_at&&v==='Unused'?'disabled':''}>${l}</option>`).join('')}</select></label></div><button type="button" id="gs-remove-device" class="gs-danger">Remove from inventory</button>`,async()=>{await change('device_update',{id,member_id:value('gs-transfer-owner')||null,...conditionData(value('gs-asset-status')),asset_tag:value('gs-asset-tag')});message('Device updated.');});$('gs-remove-device').onclick=()=>{if(confirm('Remove this device from active inventory? Purchase history will be retained.'))run(async()=>{await change('device_remove',{id});$('gs-dialog').close();message('Device removed from active inventory.');});};}
  function addClient(){
    dialog('Add client',`<div class="gs-form-grid">${field('Client / company name','gs-client-name','','required maxlength="150"')}<label>Section<select id="gs-client-section"><option value="client">Client · 1) Emails - entry</option><option value="bizdev">Biz Dev · 1) BD Emails - entry</option></select></label><label>Account manager<select id="gs-client-manager" required>${memberOptions()}</select></label><label>Default setup owner (optional)<select id="gs-client-operator">${memberOptions()}</select></label></div>`,async()=>{
      const existing=new Set(state().clients.map(c=>c.id));
      const r=await change('client_save',{name:value('gs-client-name'),section:value('gs-client-section'),account_manager_id:value('gs-client-manager'),operator_id:value('gs-client-operator')||null});
      openClient(r.state.clients.find(c=>!existing.has(c.id)).id);message('Client added. Add its first domain to get started.');
    });
  }
  function editSetup(id){
    const s=state().setups.find(s=>s.id===id)||{};
    const c=state().clients.find(c=>c.id===(s.client_id||selectedClient));
    if(!c)return addClient();
    dialog(id?'Edit domain':`Add domain to ${c.name}`,`<div class="gs-form-grid"><label>Account manager<select id="gs-job-manager" required>${memberOptions(s.account_manager_id||c.account_manager_id)}</select></label><label>Setup owner<select id="gs-job-operator" required>${memberOptions(s.operator_id||c.operator_id)}</select></label>${field('Email','gs-job-email',s.email,'type="email" required maxlength="254"')}<label>Stage<select id="gs-job-stage">${stageOptions(s.stage||'Not started',Boolean(s.purchase_at))}</select></label>${field('Email order in sheet','gs-job-order',s.order||1,'type="number" min="1" max="1000" required')}<label class="gs-wide">Device (optional)<select id="gs-job-device"><option value="">Allocate a device later</option></select></label><label class="gs-wide">Notes (optional)<textarea id="gs-job-notes" maxlength="1000">${esc(s.notes)}</textarea></label></div>`,async()=>{
      const r=await change('setup_save',{...(id?{id}:{}),client_id:c.id,stage:value('gs-job-stage'),account_manager_id:value('gs-job-manager'),operator_id:value('gs-job-operator'),section:c.section,client_name:c.name,email:value('gs-job-email'),device_id:value('gs-job-device')||null,order:Number(value('gs-job-order')),notes:value('gs-job-notes')});
      const job=r.state.setups.find(j=>j.id===id)||r.state.setups.at(-1);
      message(job.sync.status==='error'?'Domain saved; sheet sync needs attention. See its row.':'Domain saved.',job.sync.status==='error');
    });
    function options(){const owner=value('gs-job-operator');$('gs-job-device').innerHTML='<option value="">Allocate a device later</option>'+state().devices.filter(d=>d.id===s.device_id||available(d,id)&&d.member_id===owner).map(d=>`<option value="${esc(d.id)}" ${d.id===s.device_id?'selected':''}>${esc(label(d))} · ${esc(condition(d))} · ${esc(d.asset_tag)}</option>`).join('');if(s.purchase_device&&!s.stage.startsWith('Failed at'))$('gs-job-device').disabled=true;}
    $('gs-job-operator').onchange=options;options();
    $('gs-job-device').required=false;
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
  $('gs-form').onsubmit=ev=>{ev.preventDefault();run(async()=>{await submitForm();$('gs-dialog').close();});};
  for(const id of ['gs-dialog-close','gs-dialog-cancel'])$(id).onclick=()=>$('gs-dialog').close();
  $('gs-refresh').onclick=()=>run(async()=>{await refresh();message('Team tracker refreshed.');});
  $('gs-add-setup').onclick=()=>selectedClient?editSetup():addClient();
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
    if(b.dataset.sync)return run(async()=>{message('Reconciling the sheet row…');const r=await api('/setups/'+b.dataset.sync+'/sync',{});accept(r);const job=r.state.setups.find(s=>s.id===b.dataset.sync);message(job.sync.status==='synced'?'Sending email verified Live in the strategy sheet.':job.sync.message,job.sync.status!=='synced');});
    if(b.dataset.removeSetup&&confirm('Archive this setup? Device usage history will be retained.'))run(async()=>{await change('setup_remove',{id:b.dataset.removeSetup});message('Setup archived.');});
  };
  $('gs-setup-list').onchange=e=>{
    const select=e.target.closest('[data-set-stage]');if(!select)return;
    const id=select.dataset.setStage,next=select.value,job=state().setups.find(s=>s.id===id);if(next===job.stage)return;
    if(next==='Setup complete'&&!job.email){renderSetups();editSetup(id);$('gs-job-stage').value=next;return;}
    run(async()=>{try{message(next==='Setup complete'?'Saving completion and syncing the Live sheet row…':'Saving setup stage…');const r=await change('setup_stage',{id,stage:next});const updated=r.state.setups.find(s=>s.id===id);message(updated.sync.status==='error'?'Setup complete and saved. Sheet sync needs attention; see its row.':next==='Setup complete'?'Setup complete. The domain is now at the bottom of this client’s list.':'Stage updated.',updated.sync.status==='error');}finally{renderSetups();}});
  };
  $('gs-check-sheet').onclick=()=>run(async()=>{const r=await api('/check-sheet',{});$('gs-sheet-status').textContent=r.message;message(r.ready?'Sheet write access is ready.':'Sheet write access needs attention.',!r.ready);});
  // Refresh shared counts without disturbing an open edit form or active request.
  setInterval(()=>{if(!busy&&!$('gs-dialog').open&&!document.hidden)refresh().catch(err=>message(err.message,true));},30000);
  document.addEventListener('visibilitychange',()=>{if(!document.hidden&&!busy&&!$('gs-dialog').open)refresh().catch(err=>message(err.message,true));});
  run(refresh);
})();
