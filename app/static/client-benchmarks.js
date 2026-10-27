(function () {
  'use strict';
  const get = key => document.getElementById('bench-' + key);
  const stamp = value => new Date(value).toLocaleString('en-GB', {timeZone:'Europe/London',dateStyle:'medium',timeStyle:'short'});
  const month = value => new Date(value+'-01T00:00:00Z').toLocaleDateString('en-GB',{month:'short',year:'numeric',timeZone:'UTC'});
  const node = (tag, text, cls) => { const n = document.createElement(tag); n.textContent = text; if(cls) n.className=cls; return n; };
  let loading = false;
  async function load() {
    if (loading) return;
    loading = true;
    get('refresh').disabled = true;
    get('refresh').setAttribute('aria-busy','true');
    try {
      const response = await fetch('/api/client-benchmarks', {headers:{Accept:'application/json'},cache:'no-store',signal:AbortSignal.timeout(120000)});
      const data = await response.json();
      if (!response.ok) throw new Error(response.status===401?'Please sign in again.':response.status===403?'Your account does not have access to these benchmarks.':data.error || 'Benchmarks are unavailable.');
      get('groups').replaceChildren(...data.groups.map((group,index) => {
        const tr=node('tr','');
        const title=node('th','');title.scope='row';
        const indexLabel=node('span',String(index+1).padStart(2,'0'),'benchmark-index');indexLabel.setAttribute('aria-hidden','true');
        title.append(indexLabel,node('span',group.label,'benchmark-tier'));
        const size=node('td','', 'benchmark-size');size.append(node('span',group.headcount),node('small','employees'));
        const value=node('td','', 'benchmark-value');
        value.append(node('strong',group.display));
        if(group.display.length>6) value.classList.add('benchmark-value--text');
        // Keep sample limitations visible without the internal methodology detail.
        if(!group.sufficient) value.append(node('small','More campaign data needed'));
        else if(group.dominant_client_lead_share>.5) value.append(node('small','Concentrated client sample'));
        tr.append(title,size,value);return tr;
      }));
      get('note').textContent=`Rounded six-month benchmarks${data.groups.some(g=>g.estimated)?' using estimated activity and targeting inputs':''}. Indicative, not guaranteed.`;
      get('updated').textContent=`Updated ${stamp(data.synced_at)} · London time · Syncs every ${data.polling_minutes} minutes`;
      get('stale').hidden=!data.stale;
      get('stale').textContent=`Showing the last successful snapshot; data is stale. ${data.sync_error||'A fresh source sync is pending.'}`;
      get('status').textContent=`${month(data.period[0])} — ${month(data.period.at(-1))} · ${data.period.length} complete months`;
      get('message').hidden=true; get('report').hidden=false;
    } catch(error) {
      get('report').hidden=true;get('message').hidden=false;
      get('message').textContent=error.name==='TimeoutError'?'The source is taking longer than expected. Please refresh shortly.':error.message;
      get('status').textContent='Data unavailable';
    } finally {loading=false;get('refresh').disabled=false;get('refresh').removeAttribute('aria-busy');}
  }
  get('refresh').addEventListener('click',load);
  setInterval(()=>{if(!document.hidden)load();},5*60*1000);
  load();
}());
