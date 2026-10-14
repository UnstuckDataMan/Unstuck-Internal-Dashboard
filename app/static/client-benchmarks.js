(function () {
  'use strict';
  const get = key => document.getElementById('bench-' + key);
  const number = new Intl.NumberFormat('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const month = value => new Date(value + '-01T00:00:00Z').toLocaleDateString('en-GB', { month: 'short', year: 'numeric', timeZone: 'UTC' });
  const colours = ['#6ee7b7', '#5eead4', '#c4b5fd', '#a5b4fc'];
  const ranges = ['1–10 employees', '11–200 employees', '201–1,000 employees', '1,001+ employees'];
  function element(tag, text, className) {
    const node = document.createElement(tag);
    node.textContent = text;
    if (className) node.className = className;
    return node;
  }
  async function load() {
    get('refresh').disabled = true;
    get('report').hidden = true;
    get('message').hidden = false;
    get('message').textContent = 'Loading average client data…';
    get('status').textContent = 'Connecting to Google Sheets…';
    try {
      const response = await fetch('/api/client-benchmarks', {
        headers: { Accept: 'application/json' }, signal: AbortSignal.timeout(60000), cache: 'no-store'
      });
      const data = await response.json();
      if (!response.ok) throw new Error(response.status === 401 ? 'Please sign in again.' : response.status === 403 ? 'Your account does not have access to Average Client Data.' : data.error || 'Unable to load Google Sheets. Please try again.');
      const productive = data.basis === 'lead_producing_client_month';
      const active = data.basis === 'active_campaign_month';
      const billed = data.basis === 'delivery_month';
      const unit = productive ? 'Average leads in months with leads' : active ? 'leads / active campaign month' : billed ? 'leads / supplied delivery month' : 'leads / account / calendar month';
      get('groups').replaceChildren(...data.groups.map((group, index) => {
        const card = element('article', '', 'bench-group');
        card.style.setProperty('--accent', colours[index]);
        card.append(element('h2', group.label), element('p', ranges[index], 'bench-caption'),
          element('strong', group.average === null ? '—' : number.format(group.average)),
          element('p', group.average === null ? 'No recorded data' : unit, 'bench-caption'));
        return card;
      }));
      const period = `${month(data.months[0])} – ${month(data.months.at(-1))}`;
      const own = data.overall.unstuck_leads ? ' Includes Unstuck’s own prospecting.' : '';
      const basis = productive ? '' : active ? 'Includes supplied active months with zero leads.' : billed ? 'Available lead history divided by supplied lifetime delivery months; coverage may differ.' : `Includes accounts with recorded leads in each segment; all ${data.months.length} calendar months count. Active delivery months are not verified.`;
      get('note').textContent = `${data.history_scope === 'all_available' ? 'All available history · ' : ''}${period}${data.current_month_partial ? ' (current month partial, counted as one calendar month)' : ''} · ${basis}${own} Historical averages, not guaranteed volumes.`;
      get('message').hidden = true;
      get('report').hidden = false;
      get('status').textContent = 'Live Google Sheets · Refreshes on request';
    } catch (error) {
      get('message').textContent = error.name === 'TimeoutError' ? 'Google Sheets took too long to respond. Please try again.' : error.message;
      get('status').textContent = 'Data unavailable';
    } finally {
      get('refresh').disabled = false;
    }
  }
  get('refresh').addEventListener('click', load);
  load();
}());
