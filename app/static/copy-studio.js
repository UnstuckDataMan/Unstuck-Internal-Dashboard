(function () {
  'use strict';
  const get = id => document.getElementById('cs-' + id);
  if (!get('root')) return;
  const LIMITS = {email:{subject:120,initial:1200,follow_up:800,closing:500},linkedin:{connection:200,initial:600,follow_up:600,closing:400}};
  const LABELS = {subject:'Subject line',connection:'Connection note',initial:'Initial message',follow_up:'Follow-up',closing:'Closing message'};
  const state = {boot:null,client:'',epoch:0,data:null,briefRecord:null,briefDirty:false,draft:null,edited:null,dirty:false,invalid:false,busy:false,job:null,poll:null,selectedExamples:new Set(),briefEdits:new Map(),factDecisions:new Map(),resolutions:new Map()};
  let costRequest=0;
  const clone = value => JSON.parse(JSON.stringify(value));
  const asList = value => Array.isArray(value) ? value : [];
  const title = record => record?.content?.title || record?.content?.name || record?.name || record?.source_key || record?.kind || 'Untitled record';
  const friendly = value => String(value || '').replaceAll('_',' ');
  const el = (tag,text,cls) => { const node=document.createElement(tag);if(text!==undefined)node.textContent=String(text);if(cls)node.className=cls;return node; };
  const setStatus = (text,tone='') => {get('status').textContent=text;get('status').dataset.tone=tone;};
  const badge = status => {const node=el('span',friendly(status||'pending'),'cs-badge');node.dataset.status=status||'pending';return node;};
  const clientPath = (id=state.client) => '/clients/' + encodeURIComponent(id);
  const isCurrent = (client,epoch) => state.client===client && state.epoch===epoch;
  const issues = content => asList(content?.review?.issues);
  const blocking = content => issues(content).some(issue=>['error','blocker'].includes(issue.severity));
  const dateLabel = value => {if(!value)return '';const date=new Date(value);return Number.isFinite(date.getTime())?date.toLocaleString('en-GB',{dateStyle:'medium',timeStyle:'short'}):'';};
  async function request(path,body) {
    const response=await fetch('/api/copy-studio'+path,{method:body===undefined?'GET':'POST',headers:{Accept:'application/json',...(body===undefined?{}:{'Content-Type':'application/json','X-Copy-Studio-CSRF':state.boot?.csrf||''})},...(body===undefined?{}:{body:JSON.stringify(body)}),credentials:'same-origin',cache:'no-store'});
    let data;try{data=await response.json();}catch(error){throw new Error('The server returned an unreadable response. Please retry.');}
    if(!response.ok) {
      const detail=data?.error||data?.detail;
      throw new Error(response.status===401?'Please sign in again.':response.status===403?'Your account does not have access to this action. Refresh the page if your session has expired.':typeof detail==='string'?detail:Array.isArray(detail)?detail.map(item=>`${item.loc?.at(-1)||'Input'}: ${item.msg||'Check this value'}`).join(' '):detail?.message||'This action could not be completed. Please retry.');
    }
    return data;
  }
  function renderCosts(data) {
    get('costs').hidden=false;
    const money=new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',minimumFractionDigits:2,maximumFractionDigits:2});
    for(const [key,id] of [['today','today'],['month','month'],['all_time','all-time']]){
      const target=get('cost-'+id),bucket=data?.[key];target.replaceChildren();
      if(!bucket){target.textContent='Unavailable';continue;}
      if(bucket.total_calls===0){target.textContent='$0.00';continue;}
      if(Number.isFinite(bucket.amount)&&!(bucket.unpriced_calls>0)&&!(bucket.unknown_usage_calls>0)){target.textContent=money.format(bucket.amount);continue;}
      target.append(el('span',bucket.unknown_usage_calls>0?'Incomplete':'Unpriced'));
      if(Number.isFinite(bucket.known_cost)&&bucket.known_cost>0)target.append(el('small',money.format(bucket.known_cost)+' priced so far'));
    }
  }
  async function refreshCosts() {
    const token=++costRequest;get('cost-refresh').disabled=true;
    try{const result=await request('/costs');if(token===costRequest)renderCosts(result.costs||result);}
    catch(error){if(token===costRequest)renderCosts(null);}
    finally{if(token===costRequest)get('cost-refresh').disabled=false;}
  }
  function link(url,text) {
    if(typeof url!=='string')return el('span',text);
    let target;try{target=new URL(url,window.location.origin);}catch(error){return el('span',text);}
    if(!['http:','https:'].includes(target.protocol))return el('span',text);
    const a=el('a',text);a.href=target.href;a.target='_blank';a.rel='noopener noreferrer';return a;
  }
  function recordById(id) {return [...asList(state.data?.sources),...asList(state.data?.examples),...(state.data?.brief?[state.data.brief]:[])].find(record=>String(record.id)===String(id));}
  function citation(evidence) {
    const item=el('div',undefined,'cs-citation');
    if(typeof evidence==='string') {item.append(el('strong','Evidence reference'),el('p',evidence));return item;}
    const record=recordById(evidence.source_id||evidence.record_id);
    const label=record?title(record):evidence.source_id||evidence.record_id||'Source citation';
    item.append(el('strong',label+(evidence.block!==undefined?' · block '+evidence.block:'')));
    if(evidence.quote)item.append(el('p',evidence.quote));
    if(Number.isFinite(evidence.start_ms))item.append(el('small',`At ${Math.floor(evidence.start_ms/60000)}:${String(Math.floor(evidence.start_ms/1000)%60).padStart(2,'0')}`));
    const url=evidence.url||evidence.source_url||record?.content?.source_url;
    if(url)item.append(link(url,'Open source ↗'));
    return item;
  }
  function capabilities() {
    const records=asList(state.boot?.capabilities);
    get('capabilities').replaceChildren(...records.map(capability=>{
      const item=el('article',undefined,'cs-capability');
      if(typeof capability==='string'){item.append(el('strong',friendly(capability)));return item;}
      item.append(el('strong',capability.name||friendly(capability.id||capability.key)),badge(capability.status||(capability.configured?'configured':'not_configured')));
      if(capability.reason||capability.message)item.append(el('p',capability.reason||capability.message));
      return item;
    }));
    get('connection-count').textContent=records.length?`${records.filter(record=>record.configured).length} / ${records.length} configured`:'';
    if(state.boot?.model){const model=state.boot.model,item=el('article',undefined,'cs-capability');item.append(el('strong','Writing model'),badge(model.configured?'configured':'not_configured'),el('p',model.model?`${model.provider||'AI'} · ${model.model}`:'No writing model selected'),el('p',model.message||''));get('capabilities').append(item);}
  }
  function jobs() {
    const list=[...(state.job?[state.job]:[]),...asList(state.data?.jobs).filter(job=>job.id!==state.job?.id)].slice(0,6);
    get('jobs').replaceChildren(...list.map(job=>{const item=el('div',undefined,'cs-job');item.append(el('span',job.message||friendly(job.kind||job.type||'Background task')),badge(job.status));return item;}));
  }
  function renderLimits() {
    const channel=get('channel').value;
    get('limits').replaceChildren(...Object.entries(LIMITS[channel]).map(([key,value])=>{
      const wrap=el('div'),label=el('label',LABELS[key]);label.htmlFor='cs-limit-'+key;
      const input=el('input');input.id='cs-limit-'+key;input.type='number';input.min='1';input.max=String(key==='connection'?300:channel==='linkedin'?3000:10000);input.value=String(value);input.required=true;
      input.addEventListener('input',readiness);wrap.append(label,input);return wrap;
    }));
  }
  function limitValues() {return Object.fromEntries(Object.keys(LIMITS[get('channel').value]).map(key=>[key,Number(get('limit-'+key)?.value)]));}
  function limitsValid() {return Object.entries(limitValues()).every(([key,value])=>Number.isInteger(value)&&value>0&&value<=(key==='connection'?300:get('channel').value==='linkedin'?3000:10000));}
  function readiness() {
    const brief=state.data?.brief;
    const missingModel=get('mode').value==='draft'&&state.boot?.model?.configured===false;
    const missingLinkedIn=get('channel').value==='linkedin'&&!asList(state.data?.examples).some(record=>record.status==='approved'&&record.content?.source==='copy_bank'&&record.content?.channel==='linkedin');
    let reason='Approve the client brief before generating outreach.';
    if(!state.client)reason='Choose a client to begin.';
    else if(state.boot?.ready===false)reason='Generation is not configured yet. Check the connection details above.';
    else if(state.busy||state.job)reason='A client task is in progress. Your saved work remains available.';
    else if(!limitsValid())reason='Use positive character limits within the channel limits.';
    else if(missingModel)reason='Configure the writing model, or assemble approved examples without a model.';
    else if(missingLinkedIn)reason='Approve a LinkedIn Copy Bank example before drafting for this channel.';
    else if(brief?.content?.coverage?.complete===false)reason='Finish source coverage review before generating from this brief.';
    else if(brief?.status==='approved')reason=get('mode').value==='assemble'?'Select approved examples to assemble an evidence-backed sequence.':'The approved brief is ready. Set your audience and explore distinct angles.';
    const can=Boolean(state.client&&state.data&&state.boot?.ready!==false&&brief?.status==='approved'&&brief?.content?.coverage?.complete!==false&&!state.busy&&!state.job&&!missingModel&&!missingLinkedIn&&limitsValid());
    get('generation-readiness').textContent=reason;
    get('generate').disabled=!can||(get('mode').value==='assemble'&&state.selectedExamples.size===0);
    get('sync').disabled=!state.client||state.busy||Boolean(state.job);
    get('sync-selected').disabled=get('sync').disabled;
    get('extract').disabled=!state.client||!state.data||state.busy||Boolean(state.job)||state.boot?.model?.configured===false||!asList(state.data.sources).some(record=>record.kind==='transcript');
    get('save-brief').disabled=!state.briefRecord||state.busy;
    get('save-draft').disabled=!state.draft||!state.dirty||state.busy;
    const reviewRecords=[state.briefRecord,state.draft,...asList(state.data?.examples),...asList(state.data?.methodology)].filter(Boolean);
    for(const button of get('root').querySelectorAll('[data-decision]')){
      const record=reviewRecords.find(item=>String(item.id)===button.getAttribute('data-record-id'));
      if(record)button.disabled=reviewDisabled(record,record.kind==='draft',button.getAttribute('data-decision'));
    }
  }
  function renderBrief(preferred) {
    const versions=[...asList(state.data?.pending_briefs),...(state.data?.brief?[state.data.brief]:[])].sort((a,b)=>String(b.created_at||'').localeCompare(String(a.created_at||'')));
    const brief=versions.find(record=>String(record.id)===String(preferred||state.briefRecord?.id))||versions[0],content=brief?.content||{};
    state.briefRecord=brief||null;
    state.briefDirty=false;
    get('brief-version-wrap').hidden=!versions.length;
    get('brief-version').replaceChildren(...versions.map(record=>{const option=el('option',`${friendly(record.status)} · ${dateLabel(record.created_at)||title(record)}`);option.value=String(record.id);return option;}));
    if(brief)get('brief-version').value=String(brief.id);
    state.briefEdits.clear();
    state.factDecisions.clear();
    state.resolutions.clear();
    const status=brief?.status||'not_built';get('brief-status').textContent=friendly(status);get('brief-status').dataset.status=status;
    const coverage=content.coverage||state.data?.completeness||{};
    const total=Number(coverage.total_chunks||coverage.total||0),processed=Number(coverage.processed_chunks||coverage.processed||0);
    get('progress').value=total?Math.min(100,processed/total*100):coverage.complete?100:0;
    get('completeness').textContent=coverage.complete===true?'Source processing complete':brief?'Partial or unconfirmed coverage':'No brief built yet';
    get('coverage-count').textContent=total?`${processed} / ${total} chunks`:'';
    const summary=[];
    if(!brief)summary.push(el('p','Sync onboarding knowledge, then build a brief. Extracted facts stay linked to their sources for review.'));
    else {
      summary.push(el('p',`${asList(content.facts).length} sourced facts · ${asList(content.constraints).length} copy constraints`));
      if(brief.status!=='approved')summary.push(el('p',state.data?.brief?'This version is awaiting review. Generation uses the latest approved brief.':'Review this extracted brief before generating outreach.'));
      if(coverage.complete!==true)summary.push(el('p','Coverage is incomplete or needs confirmation. Review the source material before approving this brief.'));
      if(asList(content.missing_fields).length)summary.push(el('p','Still needed: '+content.missing_fields.join(', ')));
    }
    get('brief-summary').replaceChildren(...summary);
    const constraints=asList(content.constraints),list=el('ul');constraints.slice(0,4).forEach(value=>list.append(el('li',typeof value==='string'?value:value.text||JSON.stringify(value))));
    get('constraints').replaceChildren(...(constraints.length?[el('strong','Must keep in mind'),list]:[]));
    get('brief-constraints').value=constraints.map(value=>typeof value==='string'?value:value.text||'').join('\n');
    get('facts').replaceChildren(...asList(content.facts).map((fact,index)=>{
      const item=el('div',undefined,'cs-fact'),label=el('label',fact.key||fact.category||`Fact ${index+1}`,'cs-fact-meta');
      const choice=el('label',undefined,'cs-example'),enabled=el('input');enabled.type='checkbox';enabled.checked=fact.decision!=='exclude';enabled.setAttribute('data-fact-decision',String(index));
      enabled.addEventListener('change',()=>{state.factDecisions.set(index,enabled.checked?'include':'exclude');state.briefDirty=true;readiness();});choice.append(enabled,el('span','Use this fact'));
      const input=el('textarea');input.rows=3;input.value=fact.text||'';input.id='cs-fact-'+index;label.htmlFor=input.id;input.setAttribute('data-fact-index',String(index));
      input.addEventListener('input',()=>{state.briefEdits.set(index,input.value);state.briefDirty=true;readiness();});item.append(label,choice,input);
      asList(fact.evidence).forEach(evidence=>item.append(citation(evidence)));return item;
    }));
    get('contradictions').replaceChildren(...asList(content.contradictions).map((contradiction,index)=>{
      const item=el('div',undefined,'cs-issue');item.append(el('strong',`Conflicting evidence: ${contradiction.key||'Review required'}`));
      const label=el('label','Decision notes'),input=el('textarea');input.id='cs-resolution-'+index;label.htmlFor=input.id;input.rows=2;input.value=contradiction.resolution||'';input.placeholder='Exclude the outdated fact above and record why. Notes alone do not resolve conflicting active facts.';input.addEventListener('input',()=>{state.resolutions.set(index,input.value);state.briefDirty=true;readiness();});
      item.append(label,input);return item;
    }));
    get('brief-review').replaceChildren(...(brief?[reviewBox(brief)]:[]));
  }
  function renderExamples() {
    const channel=get('channel').value;
    const examples=asList(state.data?.examples).filter(record=>record.status==='approved'&&record.content?.channel===channel&&(channel!=='linkedin'||record.content?.source==='copy_bank'));
    const allowed=new Set(examples.map(record=>String(record.id)));
    state.selectedExamples=new Set([...state.selectedExamples].filter(id=>allowed.has(id)));
    get('example-count').textContent=String(examples.length);
    get('examples').replaceChildren(...(examples.length?examples.map(record=>{
      const label=el('label',undefined,'cs-example'),input=el('input'),copy=el('span');input.type='checkbox';input.value=String(record.id);input.checked=state.selectedExamples.has(input.value);input.setAttribute('data-example-id',input.value);
      input.addEventListener('change',()=>{input.checked?state.selectedExamples.add(input.value):state.selectedExamples.delete(input.value);readiness();});
      copy.append(el('strong',title(record)),el('small',record.content?.performance_proven?'Approved example with measured context':'Approved example · performance is not implied'));label.append(input,copy);return label;
    }):[el('p','No approved '+(channel==='linkedin'?'LinkedIn Copy Bank':'email')+' examples are available yet.','cs-help')]));
    get('mode-help').textContent=get('mode').value==='assemble'?'Rework approved examples for this client and audience.':'Create fresh wording grounded in the approved client brief.';
  }
  function sourceText(record) {
    const content=record.content||{};
    const blocks=asList(content.transcript||content.blocks);
    if(blocks.length)return blocks.map(block=>typeof block==='string'?block:((block.speaker?block.speaker+': ':'')+(block.text||block.content||JSON.stringify(block)))).join('\n\n');
    if(content.text||content.body)return String(content.text||content.body);
    if(content.steps)return asList(content.steps).map(step=>friendly(step.kind)+'\n'+step.body).join('\n\n');
    return JSON.stringify(content,null,2);
  }
  function renderSources() {
    const records=[...asList(state.data?.sources),...asList(state.data?.examples)];
    get('source-count').textContent=`${records.length} records`;
    get('source-library').replaceChildren(...records.map(record=>{
      const item=el('details',undefined,'cs-source-card'),summary=el('summary',title(record)),body=el('div',undefined,'cs-details-body');summary.append(badge(record.content?.coverage?.complete===false?'partial':record.status));
      item.id='cs-source-'+record.id;
      const content=record.content||{};
      body.append(el('p',`${friendly(record.kind)} · ${dateLabel(record.created_at)||'Date unavailable'}`,'cs-source-meta'));
      if(content.source_complete===false)body.append(el('p','Full source coverage is unconfirmed. Imported blocks may not represent the complete call.','cs-issue'));
      if(content.coverage?.complete===false)body.append(el('p','This source import is partial. Review its coverage before using it.','cs-issue'));
      for(const warning of asList(content.warnings))body.append(el('p',typeof warning==='string'?warning:warning.message||JSON.stringify(warning),'cs-help'));
      if(content.source_url)body.append(link(content.source_url,'Open original source ↗'));
      const preview=el('div',sourceText(record),'cs-source-preview');body.append(preview);
      const inspect=el('button','Load full source text','cs-button cs-button-secondary cs-button-small');inspect.type='button';inspect.setAttribute('data-view-record',String(record.id));
      inspect.addEventListener('click',async()=>{
        const client=state.client,epoch=state.epoch;inspect.disabled=true;
        try{const result=await request(clientPath(client)+'/records/'+encodeURIComponent(record.id));if(!isCurrent(client,epoch))return;preview.textContent=sourceText(result.record||result);inspect.textContent='Full source text loaded';}
        catch(error){if(isCurrent(client,epoch)){setStatus(error.message,'error');inspect.disabled=false;}}
      });body.append(inspect);
      if(state.boot?.can_review&&(record.kind==='transcript'||content.source==='read_ai')&&content.source_complete!==true)body.append(confirmSource(record));
      if(record.kind==='sent_copy') {
        const prepare=el('button','Prepare reusable reference','cs-button cs-button-secondary cs-button-small');prepare.type='button';prepare.setAttribute('data-prepare-reference',String(record.id));
        prepare.addEventListener('click',async()=>{const client=state.client,epoch=state.epoch;prepare.disabled=true;try{const result=await request(clientPath(client)+'/records/'+encodeURIComponent(record.id));if(isCurrent(client,epoch)){body.append(referenceEditor(result.record||result,true));prepare.hidden=true;}}catch(error){if(isCurrent(client,epoch)){prepare.disabled=false;setStatus(error.message,'error');}}});body.append(prepare);
      }else if(record.kind==='example'&&record.status!=='approved')body.append(referenceEditor(record,false));
      body.append(reviewBox(record));item.append(summary,body);return item;
    }));
    if(!records.length)get('source-library').append(el('p','No imported records yet. Sync a configured source to start.','cs-help'));
    get('performance').replaceChildren(...asList(state.data?.performance).map(record=>{
      const item=el('article',undefined,'cs-source-card'),content=record.content||record;
      item.append(el('h3',content.campaign_name||title(record)),el('p',content.basis||'Campaign comparison','cs-help'));
      item.append(el('p',content.period?.kind==='all_time_status_snapshot'?'All-time status snapshot':`Reporting period: ${content.period?.start||content.period?.from||'start unavailable'} — ${content.period?.end||content.period?.to||'end unavailable'}`,'cs-source-meta'));
      for(const variant of asList(content.variants)){
        const row=el('div',undefined,'cs-performance-row');row.append(el('strong',variant.variant||'Copy version'),badge(variant.evidence_level||'comparison_arm'));
        row.append(el('p',`${variant.sent??'—'} sent · ${variant.positive??'—'} positive replies · ${variant.reply??'—'} total replies`,'cs-help'));
        if(asList(variant.limitations).length)row.append(el('p',variant.limitations.join(' '),'cs-help'));
        for(const sourceKey of asList(variant.copy_source_keys)){
          const source=asList(state.data?.sources).find(candidate=>candidate.source_key===sourceKey);if(!source)continue;
          const open=el('button','Review exact sent copy','cs-button cs-button-secondary cs-button-small');open.type='button';open.addEventListener('click',()=>{const target=document.getElementById('cs-source-'+source.id);if(target){target.open=true;target.scrollIntoView?.({behavior:'smooth',block:'center'});}});row.append(open);
        }item.append(row);
      }
      item.append(el('p',content.interpretation||'An observed reply leader is not proof that wording caused the result.','cs-help'));return item;
    }));
    if(!asList(state.data?.performance).length)get('performance').append(el('p','No campaign learning has been imported for this client.','cs-help'));
    get('methodology').replaceChildren(...asList(state.data?.methodology).map(record=>{const item=el('details',undefined,'cs-source-card'),summary=el('summary',title(record));summary.append(badge(record.status));item.append(summary,el('div',sourceText(record),'cs-source-preview'),reviewBox(record));return item;}));
    if(state.boot?.can_review){const seed=el('button','Prepare methodology for review','cs-button cs-button-secondary cs-button-small');seed.type='button';seed.addEventListener('click',()=>mutate('/methodology/seed',{},'Writing methodology is ready for review.'));get('methodology').append(seed);}
    else if(!asList(state.data?.methodology).length)get('methodology').append(el('p','An authorised reviewer can prepare and approve the writing methodology.','cs-help'));
  }
  function referenceEditor(record,sent) {
    const wrap=el('details',undefined,'cs-form-details'),summary=el('summary',sent?'Sanitise this sent version':'Edit reusable example'),content=clone(record.content||{}),body=el('div',undefined,'cs-details-body');if(sent)wrap.open=true;
    body.append(el('p','Remove recipient-specific names, email addresses and private details. Save a new reference for review before it can guide future copy.','cs-help'));
    const subjects=asList(content.subjects).length?[...content.subjects]:[content.subject||''];const steps=asList(content.steps).length?content.steps.map(step=>({...step})):[{kind:'initial',body:content.body||content.text||''}];
    if(content.channel!=='linkedin')subjects.forEach((subject,index)=>{const label=el('label',`Subject ${index+1}`),input=el('input');input.value=subject;input.id='cs-reference-subject-'+record.id+'-'+index;label.htmlFor=input.id;input.addEventListener('input',()=>{subjects[index]=input.value;});body.append(label,input);});
    steps.forEach((step,index)=>{const label=el('label',LABELS[step.kind]||friendly(step.kind)),input=el('textarea');input.value=step.body||'';input.rows=5;input.id='cs-reference-body-'+record.id+'-'+index;label.htmlFor=input.id;input.addEventListener('input',()=>{steps[index].body=input.value;});body.append(label,input);});
    const check=el('input');check.type='checkbox';const checkLabel=el('label',undefined,'cs-example');checkLabel.append(check,el('span','I removed recipient-specific personal data and checked this is suitable for reuse.'));
    const save=el('button','Save reference for review','cs-button cs-button-secondary cs-button-small');save.type='button';save.disabled=true;check.addEventListener('change',()=>{save.disabled=!check.checked||state.busy;});
    save.addEventListener('click',()=>{
      if(save.disabled||!check.checked)return;
      if(sent)mutate('/records/'+encodeURIComponent(record.id)+'/reference',{subject:subjects[0]||'',body:steps[0]?.body||'',sanitized:true},'Reusable reference saved. An authorised reviewer can now approve it.');
      else mutate('/records',{kind:'example',source_key:record.source_key,parent_id:record.id,content:{...content,subjects:content.channel==='linkedin'?[]:subjects,steps,sanitized:true}},'A sanitised example version is ready for review.');
    });body.append(checkLabel,save);wrap.append(summary,body);return wrap;
  }
  function confirmSource(record) {
    const wrap=el('div',undefined,'cs-review-box'),label=el('label',undefined,'cs-example'),check=el('input');check.type='checkbox';check.setAttribute('data-confirm-source',String(record.id));
    label.append(check,el('span','I checked this against the original call and confirmed the transcript is complete'));
    const notesLabel=el('label','What did you check?'),notes=el('textarea');notes.rows=2;notes.id='cs-confirm-notes-'+record.id;notesLabel.htmlFor=notes.id;notes.placeholder='Describe the original call and transcript coverage you verified.';
    const button=el('button','Confirm source completeness','cs-button cs-button-secondary cs-button-small');button.type='button';button.disabled=true;button.setAttribute('data-confirm-record',String(record.id));
    const validate=()=>{button.disabled=!check.checked||notes.value.trim().length<10||state.busy;};check.addEventListener('change',validate);notes.addEventListener('input',validate);
    button.addEventListener('click',()=>{if(button.disabled)return;mutate('/transcripts/'+encodeURIComponent(record.id)+'/confirm-complete',{confirmed:true,notes:notes.value.trim()},'Transcript completeness confirmed. Rebuild the brief to use the reviewed source.');});
    wrap.append(label,notesLabel,notes,button);return wrap;
  }
  function reviewDisabled(record,isDraft,decision) {
    const content=record.content||{};
    const activeFacts=new Set(asList(content.facts).filter(fact=>fact.decision!=='exclude').map(fact=>String(fact.id)));
    const activeConflict=asList(content.contradictions).some(item=>asList(item.fact_ids).length>=2?item.fact_ids.filter(id=>activeFacts.has(String(id))).length>=2:!item.resolution);
    const briefBlocked=record.kind==='brief'&&(state.briefDirty||content.coverage?.complete===false||activeConflict);
    const referenceBlocked=record.kind==='example'&&!content.sanitized;
    return state.busy||(decision==='approved'&&(briefBlocked||referenceBlocked||(isDraft&&(state.dirty||state.invalid||content.review?.can_approve===false||blocking(content)))));
  }
  function reviewBox(record,isDraft=false) {
    const wrap=el('div',undefined,'cs-review-box');
    if(!['brief','example','draft','methodology'].includes(record.kind))return wrap;
    if(!state.boot?.can_review){wrap.append(el('p','An authorised reviewer can approve this saved version.'));return wrap;}
    if(record.status==='approved'){wrap.append(el('p','This version is approved. Editing creates a new version for review.'));return wrap;}
    const label=el('label','Review notes'),notes=el('textarea');notes.rows=2;notes.placeholder='What should change, or what did you verify?';notes.id='cs-review-notes-'+record.id;label.htmlFor=notes.id;
    const actions=el('div',undefined,'cs-review-actions');
    for(const [decision,text] of [['approved','Approve version'],['needs_changes','Request changes'],['rejected','Reject']]) {
      const button=el('button',text,'cs-button cs-button-small '+(decision==='approved'?'cs-button-primary':decision==='rejected'?'cs-button-danger':'cs-button-secondary'));button.type='button';button.setAttribute('data-decision',decision);button.setAttribute('data-record-id',String(record.id));
      button.disabled=reviewDisabled(record,isDraft,decision);
      button.addEventListener('click',()=>{
        if(button.disabled)return;
        if(decision!=='approved'&&!notes.value.trim()){notes.focus();setStatus('Add a short note so the next edit has clear direction.','error');return;}
        mutate(record.kind==='methodology'?'/methodology/'+encodeURIComponent(record.id)+'/review':'/records/'+encodeURIComponent(record.id)+'/review',{decision,notes:notes.value.trim()},'Review saved.');
      });actions.append(button);
    }
    wrap.append(label,notes,actions);
    if(record.kind==='example'&&!record.content?.sanitized)wrap.append(el('p','Save a sanitised reference version before approving this example.'));
    if(isDraft&&(state.dirty||state.invalid))wrap.append(el('p',state.dirty?'Save your edits as a new version before approval.':'Resolve message length issues before approval.'));
    return wrap;
  }
  function draftLimits() {return {...LIMITS[state.edited?.channel||'email'],...(state.edited?.provenance?.limits||{})};}
  function updateDraftValidation() {
    const limits=draftLimits();state.invalid=false;
    for(const angle of asList(state.edited?.angles)) {
      for(const subject of asList(angle.subjects))if([...String(subject)].length>(limits.subject||120))state.invalid=true;
      for(const step of asList(angle.steps))if([...String(step.body||'')].length>(limits[step.kind]||limits.initial))state.invalid=true;
    }
    get('dirty-state').textContent=state.dirty?'Unsaved edits · save a new version':'Saved version · edits create a new version';
    get('save-draft').disabled=!state.dirty||state.busy;
    get('draft-review').replaceChildren(...(state.draft?[reviewBox(state.draft,true)]:[]));
    if(state.draft&&state.edited?.mode==='draft'&&state.draft.status!=='approved'){
      const check=el('button','Check saved draft','cs-button cs-button-secondary cs-button-small');check.type='button';check.setAttribute('data-check-draft',String(state.draft.id));check.disabled=state.dirty||state.busy||Boolean(state.job);check.addEventListener('click',()=>{if(!check.disabled)startJob('/drafts/'+encodeURIComponent(state.draft.id)+'/check',{});});get('draft-review').append(check);
    }
  }
  function copyField(value,path,labelText,limit,multiline=true) {
    const wrapper=el('div',undefined,'cs-copy-field'),header=el('div',undefined,'cs-copy-label');
    const label=el('label',labelText),input=el(multiline?'textarea':'input'),count=el('div',undefined,'cs-copy-count');
    input.id='cs-copy-'+path.replaceAll('.','-');label.htmlFor=input.id;input.value=value||'';input.setAttribute('data-copy-path',path);if(multiline)input.rows=5;
    const update=()=>{const length=[...input.value].length;count.textContent=`${length} / ${limit} characters`;count.dataset.over=String(length>limit);input.setAttribute('aria-invalid',String(length>limit));};update();
    input.addEventListener('input',()=>{let target=state.edited;const parts=path.split('.');for(const part of parts.slice(0,-1))target=target[part];target[parts.at(-1)]=input.value;state.dirty=true;update();updateDraftValidation();});
    const copy=el('button','Copy','cs-copy-action');copy.type='button';copy.addEventListener('click',async()=>{try{await navigator.clipboard.writeText(input.value);setStatus('Message copied.');}catch(error){input.focus();input.select?.();setStatus('Select and copy the message from its text field.');}});
    header.append(label,copy);wrapper.append(header,input,count);return wrapper;
  }
  function renderDraft(record) {
    state.draft=record||null;state.edited=record?clone(record.content):null;state.dirty=false;state.invalid=false;
    get('draft-empty').hidden=Boolean(record);get('draft-workspace').hidden=!record;
    get('draft-status').textContent=record?friendly(record.status):'Ready when you are';get('draft-status').dataset.status=record?.status||'';
    if(!record){get('angles').replaceChildren();return;}
    get('draft-meta').replaceChildren(el('span',friendly(state.edited.channel)),el('span',dateLabel(record.created_at)),el('span',record.parent_id?'Revised version':'Original version'));
    get('review-issues').replaceChildren(...issues(state.edited).map(issue=>{const node=el('p',issue.message||issue.code,'cs-issue');node.dataset.severity=['error','blocker'].includes(issue.severity)?'blocker':'warning';return node;}));
    const limits=draftLimits();
    get('angles').replaceChildren(...asList(state.edited.angles).map((angle,index)=>{
      const card=el('article',undefined,'cs-angle'),heading=el('header',undefined,'cs-angle-heading'),body=el('div',undefined,'cs-angle-body');
      heading.append(el('div',`Angle ${String(index+1).padStart(2,'0')}`,'cs-angle-index'),el('h3',angle.name||`Angle ${index+1}`));if(angle.rationale)heading.append(el('p',angle.rationale));
      asList(angle.subjects).forEach((subject,subjectIndex)=>body.append(copyField(subject,`angles.${index}.subjects.${subjectIndex}`,`Subject ${subjectIndex+1}`,limits.subject||120,false)));
      let followups=0;asList(angle.steps).forEach((step,stepIndex)=>{const label=(LABELS[step.kind]||friendly(step.kind))+(step.kind==='follow_up'?' '+(++followups):'');body.append(copyField(step.body,`angles.${index}.steps.${stepIndex}.body`,label,limits[step.kind]||limits.initial));});card.append(heading,body);return card;
    }));
    const ids=new Set(asList(state.edited.angles).flatMap(angle=>[...asList(angle.evidence_ids),...asList(angle.steps).flatMap(step=>asList(step.evidence_ids))]));
    const facts=asList(state.data?.brief?.content?.facts);
    const evidence=asList(state.edited.evidence).map(fact=>{const item=el('div',undefined,'cs-citation');item.append(el('strong',fact.fact_id||'Supporting fact'),el('p',fact.text||''));asList(fact.citations).forEach(source=>item.append(citation(source)));return item;});
    if(!evidence.length)for(const id of ids){const fact=facts.find(item=>String(item.id)===String(id));if(fact){const item=el('div',undefined,'cs-citation');item.append(el('strong',fact.key||id),el('p',fact.text||''));asList(fact.evidence).forEach(source=>item.append(citation(source)));evidence.push(item);}else{const source=recordById(id);evidence.push(citation({source_id:id,quote:source?title(source):undefined}));}}
    get('draft-evidence').replaceChildren(...(evidence.length?evidence:[el('p','No evidence references were returned. Review the validation checks before approval.','cs-help')]));
    updateDraftValidation();
  }
  function renderVersions(preferred) {
    const drafts=[...asList(state.data?.drafts)].sort((a,b)=>String(b.created_at||'').localeCompare(String(a.created_at||'')));
    get('version-wrap').hidden=!drafts.length;
    get('version').replaceChildren(...drafts.map(record=>{const option=el('option',`${title(record)} · ${friendly(record.status)}${record.created_at?' · '+dateLabel(record.created_at):''}`);option.value=String(record.id);return option;}));
    const chosen=drafts.find(record=>String(record.id)===String(preferred))||drafts[0];if(chosen)get('version').value=String(chosen.id);renderDraft(chosen);
  }
  function renderClient(preferred) {
    get('start').hidden=true;get('workspace').hidden=false;
    renderBrief(preferred);renderExamples();renderSources();jobs();
    if(!state.dirty)renderVersions(preferred);
    readiness();
  }
  async function reload(client=state.client,epoch=state.epoch,preferred) {
    const data=await request(clientPath(client));if(!isCurrent(client,epoch))return false;state.data=data;renderClient(preferred);return true;
  }
  async function selectClient() {
    state.epoch++;const epoch=state.epoch,client=get('client').value;state.client=client;state.data=null;state.briefRecord=null;state.draft=null;state.edited=null;state.dirty=false;state.busy=false;state.job=null;state.selectedExamples.clear();
    if(state.poll)clearTimeout(state.poll);state.poll=null;get('workspace').hidden=true;get('start').hidden=false;get('jobs').replaceChildren();readiness();
    get('client-state').textContent=client?'Loading client context…':'Choose a client to begin';
    if(!client){setStatus('Choose a client to begin.');return;}
    setStatus('Loading the client’s latest brief and saved work…');
    try{if(await reload(client,epoch)){
      get('client-state').textContent='Client workspace ready';setStatus('Client context loaded.');
      const active=asList(state.data.jobs).find(job=>['queued','running'].includes(job.status));
      if(active){state.job=active;jobs();readiness();state.poll=setTimeout(()=>pollJob(active,client,epoch),1500);setStatus('A saved client task is still running. Following its progress…');}
    }}catch(error){if(isCurrent(client,epoch)){get('client-state').textContent='Could not load client';setStatus(error.message,'error');}}
  }
  async function mutate(path,body,message,preferred) {
    if(state.busy||!state.client)return;
    const client=state.client,epoch=state.epoch;state.busy=true;readiness();setStatus('Saving…');
    try{const result=await request(path.startsWith('/methodology')?path:clientPath(client)+path,body);if(!isCurrent(client,epoch))return;state.dirty=false;await reload(client,epoch,preferred||result.record?.id||result.id||state.draft?.id);if(isCurrent(client,epoch))setStatus(message,'success');}
    catch(error){if(isCurrent(client,epoch))setStatus(error.message,'error');}
    finally{if(isCurrent(client,epoch)){state.busy=false;readiness();if(state.draft)updateDraftValidation();}}
  }
  async function pollJob(job,client,epoch) {
    if(!isCurrent(client,epoch))return;
    const id=job.id||job.job_id;
    if(!id){state.job=null;readiness();setStatus('The server did not return a task identifier. Refresh the client workspace.','error');return;}
    try{
      const latest=await request(clientPath(client)+'/jobs/'+encodeURIComponent(id));if(!isCurrent(client,epoch))return;
      const current=latest.job||latest;state.job=current;jobs();
      if(['completed','complete','succeeded','success','done'].includes(current.status)){
        state.job=null;await reload(client,epoch,current.result?.record_id||current.result?.id);
        if(isCurrent(client,epoch)){
          const partial=current.result?.complete===false||current.result?.status==='partial'||current.result?.coverage?.complete===false;
          setStatus(partial?'Sync finished with partial source coverage. Review the source library before using this material.':'Your client workspace has been updated.',partial?'warning':'success');
        }refreshCosts();
      }else if(['failed','cancelled'].includes(current.status)){state.job=null;setStatus(typeof current.error==='string'?current.error:current.error?.message||current.last_error||current.message||'The task could not be completed. Please retry.','error');readiness();refreshCosts();}
      else {setStatus(current.message||'Working through the client sources…');state.poll=setTimeout(()=>pollJob(current,client,epoch),1500);}
    }catch(error){if(isCurrent(client,epoch)){state.job=null;setStatus(error.message,'error');readiness();}}
  }
  async function startJob(path,body) {
    if(state.busy||state.job||!state.client)return;
    const client=state.client,epoch=state.epoch;state.busy=true;readiness();setStatus('Starting the client task…');
    try{const result=await request(clientPath(client)+path,body);if(!isCurrent(client,epoch))return;state.job=result.job||result;jobs();setStatus('Task started. You can follow its progress in recent activity.');state.poll=setTimeout(()=>pollJob(state.job,client,epoch),1500);}
    catch(error){if(isCurrent(client,epoch))setStatus(error.message,'error');}
    finally{if(isCurrent(client,epoch)){state.busy=false;readiness();}}
  }
  get('client').addEventListener('change',selectClient);
  get('cost-refresh').addEventListener('click',refreshCosts);
  get('sync').addEventListener('click',()=>startJob('/sync',{source:'all'}));
  get('sync-selected').addEventListener('click',()=>startJob('/sync',{source:get('sync-source').value}));
  get('extract').addEventListener('click',()=>startJob('/extract-brief',{}));
  get('channel').addEventListener('change',()=>{state.selectedExamples.clear();renderLimits();renderExamples();readiness();});
  get('mode').addEventListener('change',()=>{renderExamples();readiness();});
  get('generate-form').addEventListener('submit',event=>{
    event.preventDefault();if(get('generate').disabled)return;
    startJob('/generate',{channel:get('channel').value,mode:get('mode').value,audience:{subsector:get('subsector').value.trim(),geography:get('geography').value.trim(),company_size:get('company-size').value.trim(),role:get('role').value.trim(),offer:get('offer').value.trim()},example_ids:[...state.selectedExamples],limits:limitValues()});
  });
  get('brief-version').addEventListener('change',()=>{renderBrief(get('brief-version').value);readiness();});
  get('brief-constraints').addEventListener('input',()=>{state.briefDirty=true;readiness();});
  get('version').addEventListener('change',()=>{const record=asList(state.data?.drafts).find(item=>String(item.id)===get('version').value);renderDraft(record);});
  get('save-draft').addEventListener('click',()=>{if(!state.draft||!state.dirty||state.busy)return;mutate('/drafts/'+encodeURIComponent(state.draft.id)+'/revisions',{content:clone(state.edited)},'A new draft version is ready for review.');});
  get('brief-form').addEventListener('submit',event=>{
    event.preventDefault();const brief=state.briefRecord;if(!brief||state.busy)return;
    const content=clone(brief.content);state.briefEdits.forEach((text,index)=>{content.facts[index].text=text;});content.constraints=get('brief-constraints').value.split('\n').map(text=>text.trim()).filter(Boolean);
    state.factDecisions.forEach((decision,index)=>{content.facts[index].decision=decision;});
    state.resolutions.forEach((text,index)=>{content.contradictions[index].resolution=text;});
    mutate('/records',{kind:'brief',source_key:'manual:brief-editor',parent_id:brief.id,content},'A new brief version is ready for review.');
  });
  async function bootstrap() {
    renderLimits();readiness();
    try{state.boot=await request('/bootstrap');const clients=asList(state.boot.clients);const placeholder=el('option',clients.length?'Select a client':'No clients are available');placeholder.value='';get('client').replaceChildren(placeholder,...clients.map(client=>{const option=el('option',client.name||client.client_name||client.id);option.value=String(client.id);return option;}));get('client').disabled=!clients.length;capabilities();renderCosts(state.boot.costs);setStatus(clients.length?'Select a client to open their workspace.':'No clients are available for this account.');readiness();}
    catch(error){get('client').replaceChildren(el('option','Workspace unavailable'));setStatus(error.message,'error');}
  }
  bootstrap();
}());
