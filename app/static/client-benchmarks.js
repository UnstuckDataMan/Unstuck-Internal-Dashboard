(function () {
  'use strict';
  const get = key => document.getElementById('bench-' + key);
  const number = value => value === null || value === undefined ? '—' : new Intl.NumberFormat('en-GB', {maximumFractionDigits:2}).format(value);
  const stamp = value => new Date(value).toLocaleString('en-GB', {timeZone:'Europe/London',dateStyle:'medium',timeStyle:'short'});
  const period = months => `${months[0]} – ${months.at(-1)} (${months.length} complete months)`;
  const node = (tag, text, cls) => { const n = document.createElement(tag); n.textContent = text; if(cls) n.className=cls; return n; };
  let loading = false;
  async function load() {
    if (loading) return;
    loading = true;
    get('refresh').disabled = true;
    try {
      const response = await fetch('/api/client-benchmarks', {headers:{Accept:'application/json'},cache:'no-store',signal:AbortSignal.timeout(120000)});
      const data = await response.json();
      if (!response.ok) throw new Error(response.status===401?'Please sign in again.':response.status===403?'Your account does not have access to these benchmarks.':data.error || 'Benchmarks are unavailable.');
      get('groups').replaceChildren(...data.groups.map(group => {
        const tr=node('tr',''); const title=node('th',group.label);title.scope='row';
        const value=node('td','');value.append(node('strong',group.display),node('small',group.confidence));
        if(group.months.length!==data.period.length) value.append(node('small',`${group.months.length}-month window`));
        tr.append(title,node('td',group.headcount),value);return tr;
      }));
      get('samples').replaceChildren(...data.groups.map(g=>node('p',`${g.label}: ${g.clients} clients · ${number(g.qualifying_leads)} leads / ${number(g.equivalent_months)} equivalent months. ${g.verified_live_client_months} verified and ${g.estimated_live_client_months} estimated live client-months. Allocation exposure: ${number(g.allocation_percent.outreach)}% outreach, ${number(g.allocation_percent.approved)}% approved, ${number(g.allocation_percent.estimated)}% estimated. ${period(g.months)}. Largest account: ${number(g.dominant_client_lead_share*100)}% of leads. Leave-one-client-out rates: ${g.leave_one_client_out.length?g.leave_one_client_out.map(number).join('–'):'unavailable'}.`)));
      get('comparison').replaceChildren(...data.bands.map(b=>{
        const tr=node('tr','');[b.label,number(b.preliminary),`${number(b.average)}${b.sufficient?'':' (insufficient sample)'}`,number(b.difference)].forEach(v=>tr.append(node('td',v)));return tr;
      }));
      const q=data.quality;
      get('quality').textContent=`Source checks: ${q.duplicates||0} duplicate rows; ${q.excluded_internal_eri_test||0} excluded internal/ERI/test rows; ${q.invalid_or_future||0} invalid/future rows; ${q.unknown_size||0} unknown sizes; ${q.without_stable_lead_identity||0} records without stable lead identity. ${q.unmatched_allocation_or_activity_client_months} client-months have unknown/ineligible activity or allocation; ${q.unallocated_segment_leads} sized leads could not be attributed to eligible exposure. Monthly report: ${q.reconciliation_mismatches} differences across ${q.reconciliation_compared_client_months} compared client-months (Lead + Interested totals may differ from the report’s definition).`;
      get('warnings').replaceChildren(...[...data.warnings,...data.source_metadata.notices,data.source_metadata.eri_reference].map(x=>node('li',x)));
      get('updated').textContent=`Last synced: ${stamp(data.synced_at)} · Calculated: ${stamp(data.calculated_at)} (London time). Scheduled polling every ${data.polling_minutes} minutes.`;
      get('stale').hidden=!data.stale;
      get('stale').textContent=`Showing the last successful snapshot; data is stale. ${data.sync_error||'A fresh source sync is pending.'}`;
      get('status').textContent= data.stale?'Stale snapshot':`Reporting period: ${period(data.period)}`;
      get('message').hidden=true; get('report').hidden=false;
    } catch(error) {
      // Never leave old on-screen numbers appearing fresh after a failed request.
      get('report').hidden=true;get('message').hidden=false;
      get('message').textContent=error.name==='TimeoutError'?'The source is taking too long. Please refresh shortly.':error.message;
      get('status').textContent='Data unavailable';
    } finally {loading=false;get('refresh').disabled=false;}
  }
  get('refresh').addEventListener('click',load);
  setInterval(()=>{if(!document.hidden)load();},5*60*1000);
  load();
}());
