(function () {
  'use strict';
  const get = key => document.getElementById('roi-' + key);
  const set = (key, value) => { get(key).textContent = value; };
  const form = get('form');
  const fields = ['leads', 'conversion', 'value', 'rate'].map(get);
  const currency = get('currency');
  const mode = get('mode');
  const churnEnabled = get('churn-enabled');
  const churn = get('churn');
  const number = new Intl.NumberFormat('en-GB', { maximumFractionDigits: 3 });
  const rates = { GBP: '5490', USD: '7990' };
  let previousCurrency = currency.value;
  const now = new Date();
  const london = new Intl.DateTimeFormat('en-GB', { timeZone: 'Europe/London', year: 'numeric', month: 'numeric' }).formatToParts(now);
  const year = Number(london.find(part => part.type === 'year').value);
  const month = Number(london.find(part => part.type === 'month').value) - 1;
  const dates = Array.from({ length: 12 }, (_, index) => new Date(Date.UTC(year, month + index, 1)));
  const label = date => date.toLocaleDateString('en-GB', { month: 'short', year: 'numeric', timeZone: 'UTC' });
  set('period', label(dates[0]) + ' – ' + label(dates[11]));
  function drawChart(forecast, money, code) {
    const width = 520, height = 250, left = 64, right = 14, top = 18, bottom = 40;
    const maximum = Math.max(1, ...forecast.months.map(row => row.revenue));
    const x = index => left + index / 11 * (width - left - right);
    const y = revenue => top + (1 - revenue / maximum) * (height - top - bottom);
    const compact = new Intl.NumberFormat(code === 'USD' ? 'en-US' : 'en-GB', { style: 'currency', currency: code, notation: 'compact', maximumFractionDigits: 1 });
    let svg = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Monthly revenue added by Unstuck from ${label(dates[0])} to ${label(dates[11])}. Exact figures are in the monthly table.">`;
    for (let tick = 0; tick <= 4; tick++) {
      const value = maximum * tick / 4;
      svg += `<line x1="${left}" y1="${y(value)}" x2="${width - right}" y2="${y(value)}" stroke="rgba(255,255,255,.1)"/><text x="${left - 8}" y="${y(value) + 4}" text-anchor="end" fill="#aaa6be" font-size="10">${compact.format(value)}</text>`;
    }
    svg += `<polyline points="${forecast.months.map((row, index) => `${x(index)},${y(row.revenue)}`).join(' ')}" fill="none" stroke="#6ee7b7" stroke-width="2.5" stroke-linejoin="round"/>`;
    forecast.months.forEach((row, index) => {
      svg += `<circle cx="${x(index)}" cy="${y(row.revenue)}" r="3.5" fill="#6ee7b7"><title>${label(dates[index])}: ${money.format(row.revenue)} · ${row.sales} new sales</title></circle>`;
      if (index % 2 === 0 || index === 11) svg += `<text x="${x(index)}" y="${height - 17}" text-anchor="middle" fill="#aaa6be" font-size="10">${dates[index].toLocaleDateString('en-GB', { month: 'short', timeZone: 'UTC' })}</text>`;
    });
    get('chart').innerHTML = svg + '</svg>';
    get('months').innerHTML = forecast.months.map((row, index) => `<tr><td>${label(dates[index])}</td><td>${number.format(row.sales)}</td><td>${money.format(row.revenue)}</td></tr>`).join('');
  }
  function update() {
    const code = currency.value === 'USD' ? 'USD' : 'GBP';
    const money = new Intl.NumberFormat(code === 'USD' ? 'en-US' : 'en-GB', { style: 'currency', currency: code });
    document.querySelectorAll('[data-currency-symbol]').forEach(symbol => { symbol.textContent = code === 'USD' ? '$' : '£'; });
    set('currency-code', code);
    set('value-description', mode.value === 'retainer' ? 'Monthly retainer revenue per client' : 'One-off revenue per deal');
    churnEnabled.disabled = mode.value !== 'retainer';
    const applyChurn = mode.value === 'retainer' && churnEnabled.checked;
    get('churn-controls').hidden = !applyChurn;
    const values = fields.map(field => field.value === '' ? NaN : field.valueAsNumber);
    const valid = values.map((value, index) => Number.isFinite(value) && value >= 0 && (index !== 1 || value <= 100) && (index !== 3 || value > 0));
    fields.forEach((field, index) => field.setAttribute('aria-invalid', String(field.value !== '' && !valid[index])));
    const churnValue = applyChurn ? (churn.value === '' ? NaN : churn.valueAsNumber) : 0;
    const validChurn = Number.isFinite(churnValue) && churnValue >= 0 && churnValue <= 100;
    churn.setAttribute('aria-invalid', String(applyChurn && churn.value !== '' && !validChurn));
    get('result').dataset.state = '';
    set('error', '');
    function clear(message) {
      ['ratio', 'revenue', 'net', 'deals'].forEach(key => set(key, '—'));
      set('summary', message);
      set('chart', 'Complete your inputs to see the revenue forecast.');
      set('months', '');
      set('timing', '');
    }
    if (!valid.every(Boolean) || !validChurn) {
      clear('Complete valid inputs to see your potential return.');
      const invalid = fields.findIndex((field, index) => field.value !== '' && !valid[index]);
      if (invalid !== -1) set('error', ['Enter a non-negative lead volume.', 'Enter a conversion rate between 0% and 100%.', 'Enter a non-negative conversion value.', 'Enter a monthly rate greater than zero.'][invalid]);
      else if (applyChurn && !validChurn) set('error', 'Enter annual churn between 0% and 100%.');
      return;
    }
    const [leads, conversion, value, rate] = values;
    let forecast;
    try { forecast = window.roiForecast({ leads, conversion, value, rate, mode: mode.value, churn: churnValue }); }
    catch (error) { clear('Check your figures to calculate the projection.'); set('error', error.message); return; }
    set('deals', number.format(forecast.deals));
    set('ratio', forecast.ratio.toLocaleString('en-GB', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ':1');
    set('revenue', money.format(forecast.revenue));
    set('net', money.format(forecast.net));
    get('result').dataset.state = forecast.ratio < 1 ? 'below' : 'covered';
    set('summary', forecast.ratio < 1 ? 'Projected annual revenue is below your annual Unstuck investment.' : 'Projected annual revenue covers your annual Unstuck investment.');
    set('timing', `Your assumptions imply ${number.format(forecast.expectedDeals)} sales per year. This timeline books ${number.format(forecast.deals)} whole sales as they accumulate, starting with zero existing clients. ${applyChurn ? 'Annual churn is spread across months using compound retention; each new client starts paying in their sale month.' : 'No churn is applied.'}`);
    drawChart(forecast, money, code);
  }
  currency.addEventListener('change', () => {
    rates[previousCurrency] = get('rate').value;
    previousCurrency = currency.value === 'USD' ? 'USD' : 'GBP';
    get('rate').value = rates[previousCurrency];
    update();
  });
  form.addEventListener('input', update);
  form.addEventListener('change', update);
  form.addEventListener('submit', event => event.preventDefault());
  form.addEventListener('reset', () => setTimeout(() => {
    rates.GBP = '5490'; rates.USD = '7990';
    get('rate').value = rates[currency.value === 'USD' ? 'USD' : 'GBP'];
    update();
  }, 0));
  update();
})();
