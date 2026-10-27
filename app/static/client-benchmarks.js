(function () {
  'use strict';
  const view = document.querySelector('[data-benchmark-endpoint]');
  const endpoint = view?.dataset.benchmarkEndpoint || '/api/client-benchmarks';
  const calculator = view?.dataset.benchmarkCalculator === 'true' ? document.getElementById('roi-leads') : null;
  const selection = document.getElementById('roi-benchmark-selection');
  const clearSelection = () => { if (selection) selection.textContent = ''; };
  const get = key => document.getElementById('bench-' + key);
  const scale = calculator ? get('scale') : null;
  const normalCopy = {
    headline: get('headline')?.textContent || 'Average monthly leads.',
    'intro-copy': get('intro-copy')?.textContent || 'Recent client performance, grouped by the size of companies they target.',
    metric: get('metric')?.textContent || 'Average leads/month',
    caption: get('caption')?.textContent || 'Average monthly leads by target company size',
  };
  const stamp = value => new Date(value).toLocaleString('en-GB', {timeZone:'Europe/London',dateStyle:'medium',timeStyle:'short'});
  const month = value => new Date(value+'-01T00:00:00Z').toLocaleDateString('en-GB',{month:'short',year:'numeric',timeZone:'UTC'});
  const node = (tag, text, cls) => { const n = document.createElement(tag); n.textContent = text; if(cls) n.className=cls; return n; };
  const usable = group => /^~\d+$/.test(group.display) && group.sufficient && Number.isFinite(group.average) && group.average >= 0 && Number.isSafeInteger(Number(group.display.slice(1)) * 3);
  const canScale = data => Boolean(scale && data && !data.stale && data.groups.some(usable));
  let loading = false;
  let lastData = null;
  let factor = 1;
  function updateScale() {
    const scaled = factor === 3;
    if (scale) {
      scale.disabled = loading || !canScale(lastData);
      scale.setAttribute('aria-pressed', String(scaled));
    }
    const copy = scaled ? {
      headline: 'Scaled monthly leads.',
      'intro-copy': 'Explore a 3× lead-volume scenario for your campaign.',
      metric: 'Scaled leads/month',
      caption: 'Scaled scenario: three times the benchmark lead volumes by target company size',
    } : normalCopy;
    for (const [key, text] of Object.entries(copy)) if (get(key)) get(key).textContent = text;
  }
  function render(data) {
    updateScale();
    get('groups').replaceChildren(...data.groups.map((group,index) => {
      const display = factor === 3 && usable(group) ? `~${Number(group.display.slice(1)) * factor}` : group.display;
      const tr=node('tr','');
      const title=node('th','');title.scope='row';
      const indexLabel=node('span',String(index+1).padStart(2,'0'),'benchmark-index');indexLabel.setAttribute('aria-hidden','true');
      title.append(indexLabel,node('span',group.label,'benchmark-tier'));
      const size=node('td','', 'benchmark-size');size.append(node('span',group.headcount),node('small','employees'));
      const value=node('td','', 'benchmark-value');
      value.append(node('strong',display));
      if(display.length>6) value.classList.add('benchmark-value--text');
      // Keep sample limitations visible without the internal methodology detail.
      if(!group.sufficient) value.append(node('small','More campaign data needed'));
      else if(group.dominant_client_lead_share>.5) value.append(node('small','Concentrated client sample'));
      // Only apply a published numeric benchmark; expired snapshots retain their raw rate.
      if (calculator && usable(group)) {
        const use = node('button', 'Use in forecast', 'benchmark-use');
        use.type = 'button';
        use.setAttribute('aria-label', `Use ${group.label}${factor === 3 ? ' scaled scenario' : ' benchmark'} in forecast`);
        const scaled = factor === 3;
        use.addEventListener('click', () => {
          // Apply the visible planning figure without changing source averages or the model.
          calculator.value = display.slice(1);
          calculator.dispatchEvent(new Event('input', {bubbles:true}));
          if (selection) selection.textContent = `Based on ${group.label}${scaled ? ' · Scale client 3×' : ''} · ${month(data.period[0])} — ${month(data.period.at(-1))}`;
          calculator.focus();
        });
        value.append(use);
      }
      tr.append(title,size,value);return tr;
    }));
    get('note').textContent=`${factor === 3 ? 'Scaled scenario: 3× benchmark volumes. ' : ''}Six-month benchmarks${data.groups.some(g=>g.estimated)?' using estimated activity and targeting inputs':''}. Indicative, not guaranteed.`;
    get('updated').textContent=`Updated ${stamp(data.synced_at)} · London time · Syncs every ${data.polling_minutes} minutes`;
    get('stale').hidden=!data.stale;
    get('stale').textContent=`Showing the last successful snapshot; data is stale. ${data.sync_error||'A fresh source sync is pending.'}`;
    get('status').textContent=`${month(data.period[0])} — ${month(data.period.at(-1))} · ${data.period.length} complete months`;
    get('message').hidden=true; get('report').hidden=false;
  }
  calculator?.addEventListener('input', clearSelection);
  calculator?.form?.addEventListener('reset', () => {
    clearSelection();
    factor = 1;
    if (lastData) render(lastData);
    else updateScale();
  });
  scale?.addEventListener('click', () => {
    if (loading || !canScale(lastData)) return;
    factor = factor === 1 ? 3 : 1;
    render(lastData);
  });
  async function load() {
    if (loading) return;
    loading = true;
    updateScale();
    get('refresh').disabled = true;
    get('refresh').setAttribute('aria-busy','true');
    try {
      const response = await fetch(endpoint, {headers:{Accept:'application/json'},cache:'no-store',signal:AbortSignal.timeout(120000)});
      const data = await response.json();
      if (!response.ok) throw new Error(response.status===401?'Please sign in again.':response.status===403?'Your account does not have access to these benchmarks.':data.error || 'Benchmarks are unavailable.');
      lastData = data;
      if (!canScale(data)) factor = 1;
      render(data);
    } catch(error) {
      lastData = null;
      factor = 1;
      updateScale();
      get('report').hidden=true;get('message').hidden=false;
      get('message').textContent=error.name==='TimeoutError'?'The source is taking longer than expected. Please refresh shortly.':error.message;
      get('status').textContent='Data unavailable';
    } finally {loading=false;get('refresh').disabled=false;get('refresh').removeAttribute('aria-busy');updateScale();}
  }
  get('refresh').addEventListener('click',load);
  setInterval(()=>{if(!document.hidden)load();},5*60*1000);
  load();
}());
