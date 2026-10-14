(function () {
  'use strict';
  const get = key => document.getElementById('bench-' + key);
  const set = (key, value) => { get(key).textContent = value; };
  const number = new Intl.NumberFormat('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const integer = new Intl.NumberFormat('en-GB');
  const format = value => value === null ? '—' : number.format(value);
  const month = value => new Date(value + '-01T00:00:00Z').toLocaleDateString('en-GB', { month: 'short', year: 'numeric', timeZone: 'UTC' });
  const colours = ['#6ee7b7', '#5eead4', '#c4b5fd', '#a5b4fc'];
  const ranges = ['1–10 employees', '11–200 employees', '201–1,000 employees', '1,001+ employees'];
  let data = null;
  let sourceMode = 'excel';
  function element(tag, text, className) {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    if (className) node.className = className;
    return node;
  }
  function renderSelection() {
    if (!data) return;
    const selected = data.overall;
    set('average', format(selected.average));
    set('median', format(selected.median));
    set('best', format(selected.best_month));
    set('context', selected.clients ? `${selected.label} · ${month(data.months[0])} – ${month(data.months.at(-1))}` : 'No matching leads in this period. Try another reporting period.');
    set('period-label', `${month(data.months[0])} – ${month(data.months.at(-1))}`);
    set('clients', integer.format(selected.clients));
    set('leads', integer.format(selected.leads));
    const active = data.basis === 'active_campaign_month';
    const billed = data.basis === 'delivery_month';
    set('exposure', billed ? `${selected.exposure_months} supplied delivery months` : active ? `${selected.exposure_months} active campaign-months · ${selected.campaigns} campaigns` : `${selected.clients} accounts × ${data.months.length} months = ${selected.client_months}`);
    set('scope', `${selected.unstuck_leads} of ${selected.leads} leads in this selection came from Unstuck’s own prospecting. ${active ? 'All supplied active campaign-months count, including those with zero leads. The data provider must include the complete activity roster.' : 'Includes accounts with at least one lead in the segment and period; all selected calendar months count for them. Accounts with no recorded leads are not represented.'} Results vary; these are historical observations, not guaranteed lead volumes.`);
    set('trend-label', `${selected.label} · ${active ? 'leads per active campaign (monthly sample varies)' : 'leads per account'}`);
    if (billed) {
      set('period-label', `${month(data.months[0])} – ${month(data.months.at(-1))} · available lead history`);
      set('scope', `Recorded lead history divided by supplied lifetime delivery months. ${data.matched_accounts} of ${data.roster_accounts} roster accounts matched; ${data.unmatched_accounts} unmatched accounts (${data.unmatched_delivery_months} delivery months) are not assigned to segments. ${data.unstuck_in_roster ? 'Unstuck is included with supplied delivery months.' : 'Unstuck is excluded from this metric because its delivery-month count was not supplied.'} Historical coverage may differ; this is not a complete active-campaign average.`);
    }
    const maximum = Math.max(1, ...selected.monthly.map(row => row.average || 0));
    get('months').replaceChildren(...selected.monthly.map(row => {
      const node = element('div', undefined, 'bench-month');
      const bar = element('div', undefined, 'bench-bar');
      bar.setAttribute('aria-hidden', 'true');
      const fill = element('span');
      fill.style.width = `${(row.average || 0) / maximum * 100}%`;
      bar.append(fill);
      node.append(element('span', month(row.month)), bar, element('strong', format(row.average)));
      return node;
    }));
    get('print').disabled = !selected.clients;
  }
  function render() {
    const active = data.basis === 'active_campaign_month';
    const billed = data.basis === 'delivery_month';
    const unit = billed ? 'recorded leads / supplied delivery month' : active ? 'leads / active campaign-month' : 'leads / account / calendar month';
    set('intro-scope', billed ? 'Tracked-client lead performance, using supplied delivery-month totals.' : 'Client and Unstuck campaign performance, by the size of the companies you want to reach.');
    set('accounts-label', billed ? 'Tracked accounts represented' : 'Accounts represented (including Unstuck)');
    set('unit', unit);
    set('group-unit', 'One account · One calendar month');
    set('median-label', active ? 'Median campaign average' : 'Median account average');
    set('exposure-label', active ? 'Active campaign-months' : 'Account-months in the average');
    set('basis', active ? 'Active campaign-month benchmark · Client and Unstuck campaigns · Based on the activity roster supplied in this workbook, including zero-lead active months.' : 'Calendar-month benchmark · Client and Unstuck results · This workbook does not identify active campaigns or zero-lead active months. Upload a completed activity template to calculate leads per active campaign-month.');
    document.querySelector('.bench-trend').hidden = billed;
    get('best').parentElement.hidden = billed;
    get('start').disabled = billed;
    get('end').disabled = billed;
    set('formula', active ? 'The average: total segment leads ÷ distinct active campaign/month pairs targeting that segment. Zero-lead active months count. A campaign-month counts once per group, even across multiple bands. Each lead must be allocated to one band only. The median uses each campaign’s own active-month average. Monthly rates use that month’s active campaigns, so samples can vary.' : 'The average: segment leads ÷ accounts with a recorded lead in the segment and period ÷ selected calendar months. All selected months count, including quiet months, for those accounts. The median uses account monthly averages. This is not an active-campaign rate.');
    if (billed) {
      set('basis', 'Delivery-month proxy · Available lead history ÷ supplied lifetime delivery months. Date filters do not apply. Company-size targeting is inferred from observed leads, not campaign schedules. Lead and billing histories may have different coverage.');
      set('exposure-label', 'Supplied delivery months');
      set('formula', 'For each segment: all recorded leads from matched roster clients ÷ their supplied lifetime delivery-month counts. The full tenure of each observed client is counted once per segment, not allocated according to lead share. Clients without matching lead records or known target sizes cannot be assigned. These ratios do not establish campaign activity by month or guarantee future volumes.');
    }
    const windowLabel = `${month(data.example_window[0])} – ${month(data.example_window.at(-1))}`;
    set('example-note', `Highest recorded account-month in each segment within ${windowLabel}. These are selected strong results, not average performance or expected monthly volumes. Unstuck’s own prospecting is identified where selected. Quiet months are not used to calculate an average here. Examples may come from different months and must not be added together.`);
    get('groups').replaceChildren(...data.examples.map((example, index) => {
      const card = element('article', undefined, 'bench-group bench-example');
      card.style.setProperty('--accent', colours[index]);
      card.append(element('h4', example.label), element('span', ranges[index], 'bench-caption'));
      if (example.leads === null) {
        card.append(element('strong', '—'), element('span', 'No recorded example in this window', 'bench-caption'));
      } else {
        card.append(element('span', example.account_label, 'bench-example-account'),
          element('strong', integer.format(example.leads)),
          element('span', `leads in ${month(example.month)}`, 'bench-caption'),
          element('span', 'Selected result · 1 account · 1 month', 'bench-caption'),
          element('span', `Highest of ${example.observed_account_months} lead-producing account-months across ${example.candidate_accounts} accounts`, 'bench-caption'));
      }
      return card;
    }));
    const rows = [];
    for (const [label, field, className] of [['Monthly average', 'average', 'bench-average-row'], ['Accounts represented', 'clients', 'bench-cohort-row']]) {
      const row = element('tr', undefined, className);
      const header = element('th', label); header.scope = 'row'; row.append(header);
      data.bands.forEach(band => row.append(element('td', field === 'average' ? format(band[field]) : integer.format(band[field]))));
      rows.push(row);
    }
    data.groups.forEach((group, index) => {
      const row = element('tr');
      const header = element('th', group.label); header.scope = 'row'; row.append(header);
      for (let i = 0; i < 8; i++) {
        if (i === group.bands[0]) {
          const cell = element('td', format(group.average), 'bench-segment-cell');
          cell.colSpan = group.bands.length;
          cell.style.setProperty('--accent', colours[index]);
          cell.append(element('small', `${group.clients} accounts · ${group.leads} leads`));
          row.append(cell);
        } else if (!group.bands.includes(i)) row.append(element('td'));
      }
      rows.push(row);
    });
    get('table').replaceChildren(...rows);
    const q = data.quality;
    set('quality', `Source checks (whole tab): ${q.duplicates} exact duplicate rows removed; ${q.invalid_rows} incomplete or invalid rows excluded; ${q.excluded_category || 0} records not categorised Lead excluded; ${q.unknown_size} rows with unrecognised company sizes. ${data.unmapped_in_period} otherwise valid leads in this reporting period are excluded from the size benchmarks because their company size could not be mapped.`);
    document.getElementById('bench-general-scope').hidden = billed;
    set('definition', data.lead_definition || 'Dated, named records from the Google Sheet.');
    set('fetched', `Source ${data.source_kind === 'excel' ? 'imported' : 'fetched'} ${new Date(data.fetched_at).toLocaleString('en-GB', { timeZone: 'Europe/London', dateStyle: 'medium', timeStyle: 'short' })} (London time). ${data.source_kind === 'excel' ? `Excel snapshot · ${data.source_tab} · records dated ${data.date_min} to ${data.date_max}. This is not a live feed.` : 'Source reads are cached for up to 60 seconds.'}`);
    renderSelection();
  }
  async function load(event) {
    if (event) event.preventDefault();
    if (sourceMode === 'excel' && !get('file').files.length) {
      set('message', 'Choose an Excel .xlsx workbook first.');
      get('message').hidden = false;
      return;
    }
    get('refresh').disabled = true;
    get('file').disabled = true;
    get('google').disabled = true;
    get('print').disabled = true;
    get('report').hidden = true;
    get('message').hidden = false;
    get('message').dataset.state = '';
    get('status').dataset.state = '';
    set('status', sourceMode === 'excel' ? 'Importing Excel…' : 'Connecting to source…');
    set('message', sourceMode === 'excel' ? 'Calculating from your uploaded workbook…' : 'Reading client results from the connected Google Sheet…');
    data = null;
    try {
      const params = new URLSearchParams({ start: get('start').value, end: get('end').value });
      const options = { headers: { Accept: 'application/json' }, signal: AbortSignal.timeout(60000) };
      let url = '/api/client-benchmarks?' + params;
      if (sourceMode === 'excel') {
        const body = new FormData();
        body.append('file', get('file').files[0]);
        body.append('start', get('start').value);
        body.append('end', get('end').value);
        Object.assign(options, { method: 'POST', body });
        url = '/api/client-benchmarks/upload';
      }
      const response = await fetch(url, options);
      const result = await response.json();
      if (!response.ok) throw new Error(response.status === 401 ? 'Your session has expired. Sign in again to view client benchmarks.' : response.status === 403 ? 'Your account does not have access to client benchmarks.' : result.error || 'Unable to load the source sheet.');
      data = result;
      render();
      get('message').hidden = true;
      get('report').hidden = false;
      set('status', data.source_kind === 'excel' ? `Excel snapshot · ${data.source_tab}` : 'Connected to Google Sheets');
      get('status').dataset.state = 'ready';
    } catch (error) {
      set('message', error.name === 'TimeoutError' ? 'The source took too long to respond. Please try again.' : error.message);
      get('message').dataset.state = 'error';
      set('status', 'Results unavailable');
      get('status').dataset.state = 'error';
    } finally { get('refresh').disabled = false; get('file').disabled = false; get('google').disabled = false; }
  }
  get('filters').addEventListener('submit', load);
  get('print').addEventListener('click', () => {
    const details = document.querySelector('.bench-method');
    const wasOpen = details.open;
    details.open = true;
    const restore = () => { details.open = wasOpen; window.removeEventListener('afterprint', restore); };
    window.addEventListener('afterprint', restore);
    window.print();
  });
  get('file').addEventListener('change', () => { if (get('file').files.length) { sourceMode = 'excel'; load(); } });
  get('google').addEventListener('click', () => { sourceMode = 'google'; load(); });
}());
