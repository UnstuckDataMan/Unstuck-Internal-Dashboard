/* Gmail Set — all naming/capacity decisions are verified by the server. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id), app = $('dc-app');
  if (!app) return;
  const families = JSON.parse($('dc-family-data').textContent), csrf = app.dataset.csrf;
  const defaults = {capacity:10,watch:7,near:9,sensitivity:.72,dictionary:{},themes:[],industry:'',geography:'',default_import:'add'};
  const fresh = () => ({domains:[],clusters:{},settings:structuredClone(defaults),recommendations:[],history:[],snapshots:[]});
  const clusterPages=new Map();
  let setupBatchId=null, setupPreviewDirty=false;
  let state = fresh(), analysis = null, portfolio = null, pending = null, recommendations = [], batch = 0, dirty = false, busy = false, timer, parseSequence = 0;
  const escape = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const working = () => ({domains:state.domains,clusters:state.clusters,settings:state.settings});
  function message(text, error=false) { $('dc-message').textContent = text; $('dc-message').classList.toggle('error', error); }
  function showSheetIssues(issues=[], focus=false) {
    const panel=$('dc-sheet-issues');
    panel.hidden=!issues.length;
    $('dc-sheet-issue-list').innerHTML=issues.map(r=>{
      const location=`${r.sheet} · row ${r.row}${r.cell?` · ${r.cell}`:''}`;
      // Build links from bounded location metadata, never an arbitrary response URL.
      const url=Number.isInteger(r.sheet_id)&&/^[A-G][1-9]\d*$/.test(r.cell||'')
        ?`https://docs.google.com/spreadsheets/d/1PN5Lx-5c4nbouDl-Apw9YJgrYTuoDsBKqzk5Hh6Z0po/edit#gid=${r.sheet_id}&range=${r.cell}`:null;
      return `<div class="dc-warning"><strong>${escape(location)}</strong><p>${escape(r.reason)}</p>${url?`<a href="${escape(url)}" target="_blank" rel="noopener noreferrer">Open cell ${escape(r.cell)} in Google Sheets</a>`:''}</div>`;
    }).join('');
    if(issues.length&&focus)panel.scrollIntoView({behavior:'smooth',block:'start'});
  }
  async function api(path, data, quiet=false) {
    const res = await fetch('/api/domain-cluster' + path, {method:data === undefined ? 'GET':'POST',headers:{'Content-Type':'application/json','X-Domain-CSRF':csrf},credentials:'same-origin',cache:'no-store',...(data === undefined ? {} : {body:JSON.stringify(data)})});
    if (res.status === 401) throw new Error('Your session expired. Sign in again in another tab, then refresh after copying any unsaved domains.');
    let result; try { result = await res.json(); } catch { throw new Error('Unexpected server response. Retry in a moment.'); }
    if (!res.ok) {
      if(result.detail?.sheet_issues)showSheetIssues(result.detail.sheet_issues,true);
      throw new Error(typeof result.detail === 'string' ? result.detail : result.detail?.message || result.error || 'The request could not be completed. Check the input and retry.');
    }
    return result;
  }
  async function run(fn) {
    if (busy) return;
    busy = true; app.setAttribute('aria-busy','true'); $('dc-analyse').disabled = true; $('dc-generate').disabled = true; $('dc-save').disabled = true;
    try { await fn(); } catch (err) { message(err.message,true); }
    finally { busy=false; app.removeAttribute('aria-busy'); $('dc-analyse').disabled=false; $('dc-generate').disabled=false; $('dc-save').disabled=false; }
  }
  function setDirty(value=true) { dirty=value; $('dc-save-status').textContent = value ? 'Unsaved changes · Save to Portfolio to keep them' : portfolio ? `Saved · revision ${portfolio.revision}` : 'Working analysis · not saved'; $('dc-undo').disabled = !portfolio || !state.history?.length; if($('dc-setup-apply'))$('dc-setup-apply').disabled=value||setupPreviewDirty||!state.setup_batches?.at(-1)?.preview?.items.some(i=>i.status==='ready'&&state.setup_batches.at(-1).preview.results[i.domain]?.status!=='verified'); }
  function accept(result, saved=false) {
    const previousHistory=state.history, previousSnapshots=state.snapshots, previousBatches=state.setup_batches;
    state=result.state; if(!saved&&portfolio){state.history=previousHistory;state.snapshots=previousSnapshots;state.setup_batches=previousBatches;} analysis=result.analysis; pending=null; $('dc-preview').hidden=true;
    if (result.portfolio) { portfolio=result.portfolio; $('dc-name').value=portfolio.name; }
    recommendations=[]; $('dc-plan').disabled=true; $('dc-copy-ten').disabled=true;
    render(); syncSettings(); setDirty(!saved);
  }
  const tag = value => `<span class="dc-tag ${['Full','Over Limit'].includes(value)?'high':['Watch','Near Limit','Review Required'].includes(value)?'warn':''}">${escape(value)}</span>`;
  function render() {
    if (!analysis) return;
    const s=analysis.summary, labels=[['active','Active sending domains'],['clusters','Semantic tranches'],['watch','Approaching capacity'],['full','At capacity'],['over','Over capacity'],['remaining','Existing tranche capacity'],['review','Awaiting review']];
    $('dc-stats').innerHTML=labels.map(([k,l])=>`<div class="dc-stat"><strong>${s[k]}</strong><span>${l}</span></div>`).join('');
    $('dc-rule').textContent=`${s.planned} planned · Rules ${analysis.rule_version}`;
    $('dc-distribution').innerHTML=analysis.clusters.filter(c=>c.active).map(c=>`<span style="flex:${c.active}" title="${escape(c.name)}: ${c.active} active"></span>`).join('');
    $('dc-risks').innerHTML=analysis.warnings.length ? analysis.warnings.map(w=>`<div class="dc-warning ${['Over Limit','Full'].includes(w.risk)?'high':''}"><strong>${escape(w.cluster)} · ${escape(w.risk)}</strong> — ${w.count} ${w.kind==='semantic_overlap'?'domains with related semantic evidence across assignments':`active, ${w.projected} including Planned`} / ${analysis.settings.capacity}. ${w.kind==='semantic_overlap'?'Potential semantic overlap; review related assignments.':'Review before adding domains.'}</div>`).join('') : '<div class="dc-good">No tranche capacity warnings. Continue to review uncertain classifications.</div>';
    if (s.review) $('dc-risks').innerHTML+=`<p class="dc-caption">${s.review} active or planned domains need classification review. Generic names are not assumed to share one semantic family.</p>`;
    renderClusters(); renderRecommendations();renderSetup();
    $('dc-history').innerHTML=(state.history||[]).slice().reverse().map(h=>`<p>${escape(new Date(h.at).toLocaleString('en-GB'))} · ${escape(h.action)}</p>`).join('') + (state.snapshots||[]).slice(-5).reverse().map(s=>`<p class="dc-caption">Analysis ${escape(new Date(s.at).toLocaleString('en-GB'))}: ${s.summary.active} active, ${s.summary.review} to review · ${escape(s.rule_version)}</p>`).join('') || 'No saved changes yet.';
  }
  function filteredClusters() {
    if (!analysis) return [];
    const sourceFilter=$('dc-source-filter').value, search=$('dc-search').value.toLowerCase(), status=$('dc-status').value, risk=$('dc-risk').value, confidence=$('dc-confidence').value, owner=$('dc-owner').value.toLowerCase(), platform=$('dc-platform').value.toLowerCase();
    let rows=analysis.clusters.filter(c=>!risk||c.projected_risk===risk).map(c=>({...c,domains:c.domains.filter(r=>(!sourceFilter||(sourceFilter==='manual'?!r.sheet_sources?.length:r.sheet_sources?.includes(sourceFilter)))&&(!status||r.status===status)&&(!search||`${r.domain} ${c.name} ${r.owner} ${(r.sheet_owners||[]).join(' ')}`.toLowerCase().includes(search))&&(!owner||(r.owner+' '+(r.sheet_owners||[]).join(' ')).toLowerCase().includes(owner))&&(!platform||r.platform.toLowerCase().includes(platform))&&(!confidence||(confidence==='review'?r.review_required:!r.review_required)))})).filter(c=>c.domains.length);
    const order=['Over Limit','Full','Near Limit','Watch','Healthy'];
    if ($('dc-sort').value==='size') rows.sort((a,b)=>b.projected-a.projected||a.name.localeCompare(b.name));
    else if ($('dc-sort').value==='risk') rows.sort((a,b)=>order.indexOf(a.projected_risk)-order.indexOf(b.projected_risk)||a.name.localeCompare(b.name));
    else rows.sort((a,b)=>a.name.localeCompare(b.name));
    if ($('dc-sort').value==='date') rows.forEach(c=>c.domains.sort((a,b)=>b.added_at.localeCompare(a.added_at)));
    return rows;
  }
  function renderClusters() {
    const clusters=filteredClusters();
    $('dc-clusters').innerHTML=clusters.map(c=>`<details class="dc-tranche" data-cluster="${escape(c.id)}" open><summary><strong>${escape(c.name)}</strong><span>${c.active} / ${c.capacity} active · ${c.planned} planned ${tag(c.projected_risk)}</span></summary><div class="dc-cluster-body"><div class="dc-actions"><button class="dc-small" data-action="copy-cluster" data-cluster="${escape(c.id)}">Copy domains</button><button class="dc-small" data-action="rename-cluster" data-cluster="${escape(c.id)}">Rename</button><button class="dc-small" data-action="merge-cluster" data-cluster="${escape(c.id)}">Merge into…</button></div><div class="dc-table-wrap"><table><thead><tr><th>Domain / explanation</th><th>Status / owner</th><th>Classification</th><th>Actions</th></tr></thead><tbody>${c.domains.slice((clusterPages.get(c.id)||0)*100,((clusterPages.get(c.id)||0)+1)*100).map(r=>`<tr><td><span class="dc-domain">${escape(r.domain)}</span><small>${escape(r.reason)}</small>${(r.sheet_sources||[]).map(k=>tag(k==='client'?'Client email':'Biz Dev email')).join('')}${r.sheet_at_risk?tag('Sheet: At risk'):''}${r.sheet_needs_deleting?tag('Sheet: Needs deleting'):''}${r.tags.map(t=>tag(families[t]||t)).join('')}</td><td>${escape(r.status)}<small>${escape(r.owner||(r.sheet_owners||[]).join(', ')||'No owner')} · ${escape(r.platform||'No platform')}</small></td><td>${tag(r.review_required?'Review Required':r.override||r.reviewed?'Reviewed':'Confident')}<small>${Math.round(r.confidence*100)}% rule confidence</small></td><td><button class="dc-small" data-action="edit-domain" data-domain="${escape(r.domain)}">Edit / move</button><button class="dc-small" data-action="review-domain" data-domain="${escape(r.domain)}">Mark reviewed</button></td></tr>`).join('')}</tbody></table></div>${c.domains.length>100?`<div class="dc-actions"><button class="dc-small" data-action="page-prev" data-cluster="${escape(c.id)}" ${(clusterPages.get(c.id)||0)===0?'disabled':''}>Previous 100</button><span class="dc-caption">Page ${(clusterPages.get(c.id)||0)+1} / ${Math.ceil(c.domains.length/100)}</span><button class="dc-small" data-action="page-next" data-cluster="${escape(c.id)}" ${((clusterPages.get(c.id)||0)+1)*100>=c.domains.length?'disabled':''}>Next 100</button></div>`:''}</div></details>`).join('') || '<div class="dc-empty">No domains to show. Paste an inventory above, or adjust your filters.</div>';
    $('dc-table-count').textContent=`${clusters.reduce((n,c)=>n+c.domains.length,0)} of ${state.domains.length} total domains shown. Paused and Retired domains stay in history and do not consume active capacity.`;
    renderMap(clusters);
  }
  function renderMap(clusters) {
    // Bounded accessible overview; full data always remains in the table.
    const width=Math.max(740,clusters.length*260), height=Math.max(220,...clusters.map(c=>110+Math.min(c.domains.length,20)*35));
    const coords=new Map(clusters.map((c,i)=>[c.id,130+i*260]));
    let lines='', nodes='';
    clusters.forEach((c,i)=>{ const x=130+i*260; nodes+=`<rect x="${x-110}" y="15" width="220" height="45" rx="8" fill="#e2ebfa"/><text x="${x}" y="42" text-anchor="middle">${escape(c.name)}</text>`;
      c.domains.slice(0,20).forEach((r,j)=>{const y=100+j*35; lines+=`<path d="M${x} 60V${y}" stroke="#bdcfe9"/>`;r.tags.forEach(t=>{if(coords.has(t)) lines+=`<path d="M${x} ${y}L${coords.get(t)} 60" stroke="#a29cce" stroke-dasharray="4 5" opacity=".5"/>`;});nodes+=`<a href="#dc-clusters" data-map-cluster="${escape(c.id)}"><title>${escape(r.domain+' — '+r.reason)}</title><rect x="${x-115}" y="${y-14}" width="230" height="27" rx="5" fill="white" stroke="#d8e3f0"/><text x="${x}" y="${y+4}" text-anchor="middle">${escape(r.domain)}</text></a>`;});
    });
    $('dc-map').innerHTML=`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img" aria-label="Semantic clusters and up to twenty domains per cluster">${lines}${nodes}</svg>`;
    zoomMap();
  }
  function zoomMap() { const svg=$('dc-map').querySelector('svg'); if(svg){svg.style.width=`${Number(svg.getAttribute('viewBox').split(' ')[2])*$('dc-zoom').value/100}px`;svg.style.height=`${Number(svg.getAttribute('viewBox').split(' ')[3])*$('dc-zoom').value/100}px`;} }
  function renderRecommendations() {
    $('dc-recommendations').innerHTML=recommendations.length===10?`<div class="dc-table-wrap"><table><thead><tr><th>Select</th><th>Proposed domain</th><th>Intended tranche</th><th>Current → projected</th><th>Rationale / availability</th></tr></thead><tbody>${recommendations.map((r,i)=>`<tr><td><input type="checkbox" aria-label="Select ${escape(r.domain)}" data-proposal="${i}" ${r.available===false?'disabled':'checked'}></td><td><span class="dc-domain">${escape(r.domain)}</span><button class="dc-small" data-action="copy-domain" data-value="${escape(r.domain)}">Copy</button></td><td>${escape(r.cluster_name)}</td><td>${r.current} → ${r.projected} / ${state.settings.capacity}</td><td>${escape(r.rationale)}<small>${escape(r.availability)}${r.checked_at?' · '+escape(new Date(r.checked_at).toLocaleString('en-GB')):''}${r.price_usd?' · USD '+escape(r.price_usd):''}${r.premium?' · Premium':''}</small>${r.search_url?`<a href="${escape(r.search_url)}" target="_blank" rel="noopener noreferrer">View on Porkbun ↗</a>`:''}</td></tr>`).join('')}</tbody></table></div>`:'<div class="dc-empty">Generate a batch to see ten validated .com proposals.</div>';
    $('dc-copy-ten').disabled=recommendations.length!==10; $('dc-check-porkbun').disabled=recommendations.length!==10; $('dc-plan').disabled=recommendations.length!==10;
  }
  function renderSetup() {
    if(!$('dc-setup-controls'))return;
    const current=state.setup_batches?.at(-1);
    $('dc-setup-empty').hidden=Boolean(current);$('dc-setup-controls').hidden=!current;
    if(!current)return;
    const previousNames=Object.fromEntries([...$('dc-setup-domains').querySelectorAll('input[data-report-domain]')].map(el=>[el.dataset.reportDomain,el.value]));
    const previousSelection=new Set([...$('dc-setup-domains').querySelectorAll('input[data-setup-domain]:checked')].map(el=>el.dataset.setupDomain));
    const newBatch=setupBatchId!==current.id;setupBatchId=current.id;
    if(newBatch&&current.preview){const o=current.preview.options;$('dc-setup-url').value=o.destination;$('dc-setup-name').value=o.report_user;$('dc-setup-forward').checked=o.forward;$('dc-setup-dmarc').checked=o.dmarc;$('dc-setup-code').value=o.redirect_code;$('dc-setup-www').checked=o.include_www;$('dc-setup-path').checked=o.include_path;}
    const names=current.preview?.options.report_users||{};
    $('dc-setup-domains').innerHTML='<table><thead><tr><th>Select</th><th>New domain</th><th>Known report name (optional)</th><th>Setup result</th></tr></thead><tbody>'+current.domains.map(d=>`<tr><td><input type="checkbox" data-setup-domain="${escape(d)}" aria-label="Set up ${escape(d)}" ${newBatch||previousSelection.has(d)?'checked':''}></td><td class="dc-domain">${escape(d)}</td><td><input data-report-domain="${escape(d)}" aria-label="Report mailbox name for ${escape(d)}" maxlength="64" placeholder="Use batch name (me by default)" value="${escape(previousNames[d]??names[d]??'')}"></td><td>${escape(current.preview?.results?.[d]?.status||'Not configured')}<small>${escape(current.preview?.results?.[d]?.message||'')}</small></td></tr>`).join('')+'</tbody></table>';
    const preview=current.preview;
    $('dc-setup-preview-results').innerHTML=preview?preview.items.map(item=>`<details><summary>${escape(item.domain)} · ${escape(item.status)}</summary>${item.conflicts.map(c=>`<p>${escape(c)}</p>`).join('')}${item.actions.map(a=>`<p><strong>${escape(a.kind)} / ${escape(a.host||'root')}: ${escape(a.action)}</strong><br><code>${escape(a.kind==='dmarc'?a.payload.content+' · TTL 600':a.payload.location+' · '+a.payload.redirectType)}</code></p>`).join('')}</details>`).join(''):'';
    $('dc-setup-apply').disabled=setupPreviewDirty||dirty||!preview||!preview.items.some(i=>i.status==='ready'&&preview.results[i.domain]?.status!=='verified');
  }
  function setupOptions(){
    const selected=[...$('dc-setup-domains').querySelectorAll('input[data-setup-domain]:checked')].map(el=>el.dataset.setupDomain);
    const reportUsers=Object.fromEntries([...$('dc-setup-domains').querySelectorAll('input[data-report-domain]')].filter(el=>selected.includes(el.dataset.reportDomain)&&el.value.trim()).map(el=>[el.dataset.reportDomain,el.value.trim()]));
    return {domains:selected,options:{forward:$('dc-setup-forward').checked,destination:$('dc-setup-url').value.trim(),redirect_code:$('dc-setup-code').value,include_path:$('dc-setup-path').checked,include_www:$('dc-setup-www').checked,dmarc:$('dc-setup-dmarc').checked,report_user:$('dc-setup-name').value.trim()||'me',report_users:reportUsers}};
  }
  if($('dc-setup-preview')){
    $('dc-setup-controls').addEventListener('input',()=>{setupPreviewDirty=true;$('dc-setup-apply').disabled=true;$('dc-setup-progress').textContent='Settings changed · refresh the preview.';});
    $('dc-setup-preview').onclick=()=>run(async()=>{
      if(!portfolio||dirty)throw new Error('Save your portfolio before previewing domain setup.');
      const current=state.setup_batches?.at(-1);if(!current)throw new Error('Save a new Planned batch first.');
      const selected=setupOptions();if(!selected.domains.length)throw new Error('Select at least one new domain.');
      $('dc-setup-progress').textContent='Checking ownership, API permissions, nameservers and existing records…';
      const result=await api('/portfolios/'+portfolio.id+'/setup/preview',{batch_id:current.id,revision:portfolio.revision,...selected});
      setupPreviewDirty=false;accept(result,true);$('dc-setup-progress').textContent='Preview ready. Only ready domains will be configured.';
    });
    $('dc-setup-apply').onclick=()=>run(async()=>{
      if(dirty||setupPreviewDirty)throw new Error('Save changes and create a fresh preview first.');
      const current=state.setup_batches?.at(-1), preview=current?.preview;if(!preview)throw new Error('Preview this batch first.');
      const selected=new Set([...$('dc-setup-domains').querySelectorAll('input[data-setup-domain]:checked')].map(el=>el.dataset.setupDomain));
      const ready=preview.items.filter(i=>i.status==='ready'&&selected.has(i.domain)&&preview.results[i.domain]?.status!=='verified');
      for(let index=0;index<ready.length;index++){
        const item=ready[index];$('dc-setup-progress').textContent=`Configuring ${index+1}/${ready.length}: ${item.domain}…`;
        const result=await api('/portfolios/'+portfolio.id+'/setup/apply',{batch_id:current.id,preview_id:preview.id,revision:portfolio.revision,domain:item.domain});
        accept(result,true);
      }
      $('dc-setup-progress').textContent='Batch finished. Review each domain result below; DNS propagation may still be pending.';
    });
  }
  $('dc-check-setup').onclick=()=>run(async()=>{message('Checking server connections…');const result=await api('/integrations/check',{});if(result.sheets.issues)showSheetIssues(result.sheets.issues,true);else if(result.sheets.ready)showSheetIssues();$('dc-integration-results').innerHTML=`<section class="dc-card"><strong>Setup check</strong><p>Porkbun credentials: ${result.porkbun.credentials_valid?'Verified':'Needs attention'} · Availability reads: ${result.porkbun.read?'Allowed':'Not verified'} · DNS / forwarding: ${result.porkbun.dns?'Allowed':'Needs DNS permission'}</p><p>Portfolio storage: ${result.storage.ready?'Ready':'Needs attention'} · Live strategy sheet: ${result.sheets.ready?'Ready':'Needs attention'}</p>${[result.porkbun,result.storage,result.sheets].filter(r=>r.message).map(r=>`<p class="dc-caption">${escape(r.message)}</p>`).join('')}</section>`;message('Setup check complete.');});
  function syncSettings() {
    const s=state.settings; for(const key of ['capacity','watch','near','sensitivity']) $('dc-'+key).value=s[key];
    $('dc-dictionary').value=JSON.stringify(s.dictionary,null,2); $('dc-theme').value=s.themes[0]||''; $('dc-default-import').value=s.default_import;
  }
  async function refreshPortfolios() {
    const result=await api('/portfolios');
    $('dc-portfolios').innerHTML='<option value="">New portfolio</option>'+result.portfolios.map(p=>`<option value="${escape(p.id)}">${escape(p.name)}</option>`).join('');
    $('dc-portfolios').value=portfolio?.id||'';
  }
  async function save() {
    if(pending)throw new Error('Apply or cancel the import preview before saving.');
    const proposalsBeforeSave=recommendations;
    const result=await api('/portfolios',{state:working(),portfolio_id:portfolio?.id||null,name:$('dc-name').value.trim()||'My domain portfolio',revision:portfolio?.revision||0,action:'Save working portfolio'});
    accept(result,true); recommendations=proposalsBeforeSave;renderRecommendations();await refreshPortfolios(); message('Portfolio saved. Your classifications, settings and history are available after signing back in.');
    try {localStorage.setItem('dc-recent',portfolio.id);} catch {}  // Identifier only; never store mailbox input or private inventories in browser storage.
  }
  async function reanalyse() { const result=await api('/analyse',{state:working()}); analysis=result.analysis; render(); }
  async function change(fn, notice) { if(pending)throw new Error('Apply or cancel the import preview before making other changes.'); const next=structuredClone(state); fn(next); const result=await api('/analyse',{state:{domains:next.domains,clusters:next.clusters,settings:next.settings}}); accept(result); message(notice+' Save to Portfolio to keep this change.'); }
  async function copy(text) { try { await navigator.clipboard.writeText(text); message('Copied to clipboard.'); } catch { showDialog('Copy domains',`<label>Copy this text<textarea id="dc-copy-fallback" rows="8" readonly>${escape(text)}</textarea></label>`,async()=>{}); $('dc-copy-fallback').select(); } }
  function showDialog(title, html, onSave) {
    $('dc-dialog-title').textContent=title; $('dc-dialog-body').innerHTML=html; $('dc-dialog-error').textContent=''; $('dc-dialog').showModal();
    $('dc-dialog-save').onclick=async ev=>{ev.preventDefault();$('dc-dialog-save').disabled=true;try{await onSave();$('dc-dialog').close();}catch(err){$('dc-dialog-error').textContent=err.message;}finally{$('dc-dialog-save').disabled=false;}};
  }
  const clusterOptions=(selected='')=>'<option value="">Automatic classification</option>'+Object.entries({...families,review:'Review Required',...Object.fromEntries(Object.entries(state.clusters).map(([k,v])=>[k,v.name]))}).map(([k,v])=>`<option value="${escape(k)}" ${k===selected?'selected':''}>${escape(v)}</option>`).join('');
  $('dc-clusters').addEventListener('click',ev=>{const btn=ev.target.closest('button[data-action]');if(!btn)return;run(async()=>{
    const id=btn.dataset.cluster;
    if(btn.dataset.action.startsWith('page-')){clusterPages.set(id,Math.max(0,(clusterPages.get(id)||0)+(btn.dataset.action==='page-next'?1:-1)));renderClusters();return;}
    const domain=btn.dataset.domain, row=state.domains.find(r=>r.domain===domain);
    if(btn.dataset.action==='copy-cluster') return copy(analysis.domains.filter(r=>r.cluster===id).map(r=>r.domain).join('\n'));
    if(btn.dataset.action==='review-domain') return change(s=>{s.domains.find(r=>r.domain===domain).reviewed=true;},'Classification marked reviewed.');
    if(btn.dataset.action==='edit-domain') showDialog('Edit '+domain,`<label>Primary tranche<select id="dc-edit-cluster">${clusterOptions(row.override||'')}</select></label><label>Status<select id="dc-edit-status">${['Active','Paused','Planned','Retired'].map(v=>`<option ${v===row.status?'selected':''}>${v}</option>`).join('')}</select></label><label>Team member / owner<input id="dc-edit-owner" maxlength="100" value="${escape(row.owner)}"></label><label>Sending platform<input id="dc-edit-platform" maxlength="100" value="${escape(row.platform)}"></label><label>Internal notes<textarea id="dc-edit-notes" maxlength="1000">${escape(row.notes)}</textarea></label><p class="dc-caption">Choose Automatic classification to restore rules. Manual assignments may exceed capacity; warnings stay visible and recommendations still check the natural semantic evidence.</p>`,async()=>change(s=>{Object.assign(s.domains.find(r=>r.domain===domain),{override:$('dc-edit-cluster').value||null,reviewed:Boolean($('dc-edit-cluster').value),status:$('dc-edit-status').value,owner:$('dc-edit-owner').value,platform:$('dc-edit-platform').value,notes:$('dc-edit-notes').value});},'Domain updated.'));
    if(btn.dataset.action==='rename-cluster') showDialog('Rename tranche',`<label>Name<input id="dc-cluster-name" value="${escape(analysis.clusters.find(c=>c.id===id).name)}" maxlength="100"></label>`,async()=>{const name=$('dc-cluster-name').value.trim();if(!name)throw new Error('Enter a name.');await change(s=>{s.clusters[id]={name};},'Tranche renamed.');});
    if(btn.dataset.action==='merge-cluster') showDialog('Merge tranche',`<label>Move all members into<select id="dc-merge-target">${clusterOptions().replace('<option value="">Automatic classification</option>','')}</select></label><p class="dc-caption">All members receive a manual assignment. Over-capacity warnings remain visible.</p>`,async()=>{const target=$('dc-merge-target').value;if(target===id)throw new Error('Select a different tranche.');const members=new Set(analysis.domains.filter(r=>r.cluster===id).map(r=>r.domain));await change(s=>{s.domains.forEach(r=>{if(members.has(r.domain)){r.override=target;r.reviewed=true;}});},'Tranches merged.');});
  });});
  async function parseInput(preview=false) {
    const text=$('dc-text').value, seq=++parseSequence;
    if(!text.trim()){ $('dc-parse-status').textContent='Paste domains to begin.';return; }
    const result=await api('/parse',{text,mode:$('dc-mode').value,state:working()});
    if(seq!==parseSequence) return;
    $('dc-parse-status').textContent=`${result.parsed.unique} unique valid domains · ${result.parsed.duplicates} duplicates · ${result.parsed.invalid.length} invalid entries`;
    $('dc-invalid-list').innerHTML=result.parsed.invalid.slice(0,100).map(r=>`<p>Line ${r.line}: <code>${escape(r.input)}</code> — ${escape(r.reason)}</p>`).join('')+(result.parsed.invalid.length>100?'<p>First 100 invalid entries shown. Correct these and reanalyse to continue reviewing.</p>':'');
    if(preview){
      if(!result.parsed.unique)throw new Error('No valid domains found. Review validation feedback.');
      pending=result; $('dc-preview').hidden=false;
      const p=result.preview;
      $('dc-preview-body').innerHTML=[['Newly added',p.added],['Already existing',p.existing],['Missing active domains',p.missing],['Reactivated by replacement',p.reactivated]].map(([label,rows])=>`<details><summary>${label}: ${rows.length}</summary><p class="dc-domain">${rows.slice(0,150).map(escape).join(', ')}${rows.length>150?' …':''}</p></details>`).join('');
      $('dc-preview-note').textContent=p.mode==='replace'?'Replacement will set missing Active domains to Paused and included domains to Active. No records are deleted. Review these changes, then apply and save.':'New domains will be added as Active. Existing records and statuses are preserved.';
      $('dc-preview').scrollIntoView({behavior:'smooth',block:'nearest'}); message('Review the import before applying it. Nothing has been saved.');
    }
  }
  $('dc-text').addEventListener('input',()=>{clearTimeout(timer);pending=null;$('dc-preview').hidden=true;parseSequence++;timer=setTimeout(()=>parseInput().catch(err=>message(err.message,true)),650);});
  $('dc-analyse').onclick=()=>run(async()=>{
    if($('dc-source').value==='manual')return parseInput(true);
    message('Reading current Live checkboxes from both strategy tabs…');
    const result=await api('/strategy-sheet',{state:working()});
    pending=result;showSource(result.source);$('dc-preview').hidden=false;
    const p=result.preview;
    $('dc-preview-body').innerHTML=[['Live domains to add',p.added],['Existing live domains',p.existing],['Previously imported domains no longer Live',p.missing],['Domains becoming Active',p.reactivated]].map(([label,rows])=>`<details><summary>${label}: ${rows.length}</summary><p class="dc-domain">${rows.slice(0,150).map(escape).join(', ')}</p></details>`).join('');
    $('dc-preview-note').textContent='Live sheet rows become Active. Previously imported sheet domains no longer Live become Paused; manual-only domains keep their status. No records are deleted. Apply this preview, then save.';
    $('dc-invalid-list').textContent='';
    message(`${result.parsed.unique} unique Live domains found in a fresh scan. Review the preview before applying.${result.source.invalid.length?' Some Live rows need correction; generation will be blocked.':''}`);
    if(result.source.invalid.length)$('dc-sheet-issues').scrollIntoView({behavior:'smooth',block:'start'});else $('dc-preview').scrollIntoView({behavior:'smooth',block:'nearest'});
  });
  function showSource(source){showSheetIssues(source.invalid);$('dc-live-status').textContent=`Fresh scan ${new Date(source.checked_at).toLocaleString('en-GB')} · `+source.tabs.map(t=>`${t.kind==='client'?'Client':'Biz Dev'}: ${t.live_rows} Live rows, ${t.unique_domains} domains`).join(' · ')+`. ${source.tabs.reduce((n,t)=>n+t.at_risk_rows,0)} Live rows flagged At risk; ${source.tabs.reduce((n,t)=>n+t.needs_deleting_rows,0)} flagged Needs deleting. Live controls inclusion; these flags remain visible separately.`;}
  $('dc-source').onchange=()=>{pending=null;$('dc-preview').hidden=true;const manual=$('dc-source').value==='manual';$('dc-text').disabled=!manual;$('dc-file').disabled=!manual;$('dc-mode').disabled=!manual;message(manual?'Manual inventory mode. Paste or upload domains to analyse.':'Live mode: each Analyse or Generate fetches the latest sheet data.');};
  $('dc-apply').onclick=()=>run(async()=>{if(!pending)return;accept(pending);pending=null;$('dc-preview').hidden=true;message('Analysis ready. Review Current Tranches and risk warnings below. Save to Portfolio to keep this inventory.');});
  $('dc-cancel-import').onclick=()=>{pending=null;$('dc-preview').hidden=true;};
  $('dc-reset').onclick=()=>{$('dc-text').value='';$('dc-invalid-list').textContent='';$('dc-parse-status').textContent='Input cleared. Portfolio unchanged.';pending=null;$('dc-preview').hidden=true;parseSequence++;};
  $('dc-mode').onchange=()=>{pending=null;$('dc-preview').hidden=true;};
  $('dc-file').onchange=()=>run(async()=>{const file=$('dc-file').files[0];if(!file)return;if(!/\.(csv|txt)$/i.test(file.name)||file.size>1000000)throw new Error('Choose a CSV or TXT file up to 1 MB.');$('dc-text').value=await file.text();await parseInput();$('dc-file').value='';});
  $('dc-save').onclick=()=>run(save);
  $('dc-portfolios').onchange=()=>{const selected=$('dc-portfolios').value; if(dirty){$('dc-portfolios').value=portfolio?.id||'';showDialog('Switch portfolio?', '<p>Unsaved changes will be discarded. Save to Portfolio first if you want to keep them.</p>',async()=>loadPortfolio(selected));}else run(()=>loadPortfolio(selected));};
  async function loadPortfolio(id){if(!id){portfolio=null;state=fresh();analysis=(await api('/analyse',{state:working()})).analysis;render();syncSettings();setDirty(false);$('dc-portfolios').value='';$('dc-name').value='My domain portfolio';return;}accept(await api('/portfolios/'+encodeURIComponent(id)),true);$('dc-portfolios').value=id;try {localStorage.setItem('dc-recent',id);} catch {} message('Saved portfolio loaded.');}
  $('dc-undo').onclick=()=>run(async()=>{if(dirty)throw new Error('Save or reload your working changes before undoing a saved update.');const result=await api('/portfolios/'+portfolio.id+'/update',{revision:portfolio.revision,action:'undo'});accept(result,true);message('Last saved update undone. No domain records were deleted by replacement imports.');});
  $('dc-reanalyse').onclick=()=>{if($('dc-source').value==='sheet'){$('dc-analyse').click();return;}run(async()=>{await reanalyse();message('Analysis refreshed. User-confirmed assignments were preserved.');});};
  $('dc-new-cluster').onclick=()=>showDialog('Create semantic tranche','<label>Tranche name<input id="dc-cluster-name" maxlength="100" placeholder="A distinct commercial intent"></label><p class="dc-caption">Move domains with Edit / move. Natural semantic exposure still applies to recommendations.</p>',async()=>{const name=$('dc-cluster-name').value.trim();if(!name)throw new Error('Enter a name.');await change(s=>{s.clusters['manual-'+crypto.randomUUID().slice(0,12)]={name};},'Tranche created.');});
  $('dc-generate').onclick=()=>run(async()=>{if(pending)throw new Error('Apply or cancel the import preview before generating.');message('Checking naming rules, semantic overlap and projected capacity…');const r=await api('/generate',{state:working(),batch:batch++,require_available:$('dc-require-available').checked,live_sheet:$('dc-source').value==='sheet'});if(r.source){const oldInventory=JSON.stringify(state.domains.map(d=>[d.domain,d.status,d.override]));const newInventory=JSON.stringify(r.state.domains.map(d=>[d.domain,d.status,d.override]));const wasDirty=dirty;accept(r);showSource(r.source);setDirty(wasDirty||oldInventory!==newInventory);}recommendations=r.recommendations;renderRecommendations();message(r.shortfall?`${r.message} ${r.qualifying_count} candidates qualified; shortfall ${r.shortfall}.`:$('dc-require-available').checked?'Ten candidates verified available on Porkbun. Check timestamps and confirm at checkout.':'Ten candidates validated. Availability not checked.',Boolean(r.shortfall));});
  $('dc-recommendations').onclick=ev=>{const b=ev.target.closest('button[data-action="copy-domain"]');if(b)run(()=>copy(b.dataset.value));};
  $('dc-check-porkbun').onclick=()=>run(async()=>{message('Checking the ten proposed names on Porkbun…');const r=await api('/availability',{domains:recommendations.map(c=>c.domain)});recommendations=recommendations.map(c=>({...c,...r.results[c.domain]}));renderRecommendations();message(`${recommendations.filter(c=>c.available===true).length} available; ${recommendations.filter(c=>c.available===false).length} unavailable; ${recommendations.filter(c=>c.available==null).length} unverified. Use verified-only search to find a complete available batch.`);});
  $('dc-copy-ten').onclick=()=>run(()=>copy(recommendations.map(r=>r.domain).join('\n')));
  $('dc-plan').onclick=()=>run(async()=>{
    if(dirty||!portfolio)throw new Error('Save your portfolio before saving planned domains so capacity is checked against the latest saved inventory.');
    const selected=[...$('dc-recommendations').querySelectorAll('input[data-proposal]:checked')].map(el=>recommendations[Number(el.dataset.proposal)]);if(!selected.length)throw new Error('Select at least one domain.');
    accept(await api('/portfolios/'+portfolio.id+'/update',{revision:portfolio.revision,action:'planned',candidates:selected}),true);message(`${selected.length} domains saved as Planned. They now count towards projected capacity.`);
  });
  $('dc-settings').onclick=()=>run(async()=>{let dict;try{dict=JSON.parse($('dc-dictionary').value);}catch{throw new Error('Enter valid JSON for the dictionary.');}await change(s=>{s.settings={...s.settings,capacity:Number($('dc-capacity').value),watch:Number($('dc-watch').value),near:Number($('dc-near').value),sensitivity:Number($('dc-sensitivity').value),dictionary:dict,themes:$('dc-theme').value?[$('dc-theme').value]:[],default_import:$('dc-default-import').value};},'Settings applied.');$('dc-mode').value=state.settings.default_import;});
  for(const id of ['dc-search','dc-status','dc-source-filter','dc-risk','dc-confidence','dc-owner','dc-platform','dc-sort']) $(id).addEventListener('input',()=>{clusterPages.clear();renderClusters();});
  $('dc-zoom').oninput=zoomMap;
  $('dc-map').onclick=ev=>{const a=ev.target.closest('[data-map-cluster]');if(a){const c=[...$('dc-clusters').querySelectorAll('details')].find(c=>c.dataset.cluster===a.dataset.mapCluster);if(c){c.open=true;c.scrollIntoView({behavior:'smooth'});}}};
  function download(text,ext,type){const url=URL.createObjectURL(new Blob([text],{type}));const a=document.createElement('a');a.href=url;a.download='domain-cluster-'+$('dc-export-scope').value+'.'+ext;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
  const csvCell=v=>{let s=String(v??'');if(/^[=+@\-\t\r]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';};
  function exportData(){
    const scope=$('dc-export-scope').value,format=$('dc-export-format').value;
    if(scope==='recommendations'&&recommendations.length!==10)throw new Error('Generate ten recommendations before exporting them.');
    const rows=scope==='risks'?analysis.warnings:scope==='recommendations'?recommendations:scope==='domains'?analysis.domains.filter(r=>r.status==='Active'):analysis.domains;
    if(format==='json')return download(JSON.stringify(scope==='report'?{...analysis,recommendations,generated_at:new Date().toISOString()}:rows,null,2),'json','application/json');
    if(format==='csv'&&scope==='report'){
      const keys=['section','domain','cluster','status','active','planned','risk','details'];
      const reportRows=[['Overview','','','',analysis.summary.active,analysis.summary.planned,'',`Rules ${analysis.rule_version}; capacity ${state.settings.capacity}; ${analysis.summary.review} review required`],...analysis.clusters.map(c=>['Tranche','',c.name,'',c.active,c.planned,c.projected_risk,'']),...analysis.domains.map(r=>['Domain',r.domain,r.cluster,r.status,'','','',r.reason]),...analysis.warnings.map(w=>['Warning','',w.cluster,'',w.count,w.projected-w.count,w.risk,w.kind]),...recommendations.map(r=>['Recommendation',r.domain,r.cluster_name,'Proposed','','','',r.availability+'; '+r.rationale])];
      return download([keys,...reportRows].map(row=>row.map(csvCell).join(',')).join('\r\n'),'csv','text/csv;charset=utf-8');
    }
    if(format==='csv'){const keys=scope==='risks'?['cluster','count','projected','risk','kind']:scope==='recommendations'?['domain','cluster_name','current','projected','rationale','availability']:['domain','cluster','status','owner','platform','confidence','reason','added_at'];return download([keys.map(csvCell).join(','),...rows.map(r=>keys.map(k=>csvCell(r[k])).join(','))].join('\r\n'),'csv','text/csv;charset=utf-8');}
    let text=scope==='recommendations'?rows.map(r=>r.domain).join('\n'):scope==='domains'?rows.map(r=>r.domain).join('\n'):scope==='risks'?rows.map(r=>`${r.cluster}: ${r.risk} (${r.projected}/${state.settings.capacity})`).join('\n'):analysis.clusters.map(c=>`${c.name} — ${c.active}/${c.capacity} Active, ${c.planned} Planned — ${c.projected_risk}\n${c.domains.map(r=>r.domain+' ['+r.status+']').join('\n')}`).join('\n\n');
    if(scope==='report')text='Domain Cluster Builder\nPlanning policy; semantic similarity does not prove provider association.\nRules '+analysis.rule_version+'\n\nCurrent Tranches\n'+text+'\n\nCluster Risk Warning\n'+(analysis.warnings.map(w=>w.cluster+': '+w.risk).join('\n')||'No capacity warnings.')+'\n\nRecommended New Domains (10)\n'+(recommendations.length===10?recommendations.map(r=>r.domain).join('\n'):'Not generated.');
    download(text,'txt','text/plain;charset=utf-8');
  }
  $('dc-export').onclick=()=>run(async()=>exportData());
  $('dc-print').onclick=()=>{document.querySelectorAll('.dc-tranche').forEach(d=>d.open=true);window.print();};
  $('dc-dns').onclick=()=>run(async()=>{if(!portfolio||dirty)throw new Error('Save your current portfolio before running DNS checks.');$('dc-dns-results').textContent='Looking up public records…';const r=await api('/portfolios/'+portfolio.id+'/dns',{selector:$('dc-dkim').value});$('dc-dns-results').innerHTML=(r.commonality.length?'<p><strong>Shared DNS provider detected</strong></p>'+r.commonality.map(c=>`<p class="dc-caption">${escape(c.type)} ${escape(c.value)}: ${c.domains.length} domains. Commonality only; no risk inference.</p>`).join(''):'')+r.results.map(row=>`<details><summary>${escape(row.domain)} ${row.cached?'· cached':''}</summary><p class="dc-caption">Checked ${escape(row.checked_at)}</p>${Object.entries(row.records).map(([type,v])=>`<p><strong>${escape(type)}</strong>: ${escape(v.values?.join(' | ')||v.error)}</p>`).join('')}</details>`).join('')+`<p class="dc-caption">${escape(r.note)}</p>`;});
  window.addEventListener('beforeunload',ev=>{if(dirty){ev.preventDefault();ev.returnValue='';}});
  run(async()=>{analysis=(await api('/analyse',{state:working()})).analysis;render();syncSettings();setDirty(false);$('dc-text').disabled=true;$('dc-file').disabled=true;$('dc-mode').disabled=true;const integrations=await api('/integrations');$('dc-porkbun-status').textContent=integrations.porkbun?'Porkbun connected':'Porkbun not connected · ask an admin to enable availability checks';try{await refreshPortfolios();let recent;try {recent=localStorage.getItem('dc-recent');} catch {} if(recent&&[...$('dc-portfolios').options].some(o=>o.value===recent))await loadPortfolio(recent);}catch(err){message(err.message,true);}});
})();
