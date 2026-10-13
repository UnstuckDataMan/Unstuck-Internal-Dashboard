(function () {
  'use strict';
  const get = key => document.getElementById('roi-' + key);
  const set = (key, value) => { get(key).textContent = value; };
  const pdf = get('pdf');
  let exportSnapshot = null;
  const form = get('form');
  const scenario = get('scenario');
  const scenarioNames = { conservative: 'Conservative', baseline: 'Baseline', optimistic: 'Optimistic' };
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
  const dates = Array.from({ length: 60 }, (_, index) => new Date(Date.UTC(year, month + index, 1)));
  const label = date => date.toLocaleDateString('en-GB', { month: 'short', year: 'numeric', timeZone: 'UTC' });
  set('period', label(dates[0]) + ' – ' + label(dates[59]));
  function drawChart(forecast, money, code) {
    const width = 520, height = 250, left = 64, right = 14, top = 18, bottom = 40;
    const maximum = Math.max(1, ...forecast.timeline.map(row => row.revenue));
    const x = index => left + index / 59 * (width - left - right);
    const y = revenue => top + (1 - revenue / maximum) * (height - top - bottom);
    const compact = new Intl.NumberFormat(code === 'USD' ? 'en-US' : 'en-GB', { style: 'currency', currency: code, notation: 'compact', maximumFractionDigits: 1 });
    let svg = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Monthly revenue added by Unstuck from ${label(dates[0])} to ${label(dates[59])}. Exact figures are in the monthly table.">`;
    for (let tick = 0; tick <= 4; tick++) {
      const value = maximum * tick / 4;
      svg += `<line x1="${left}" y1="${y(value)}" x2="${width - right}" y2="${y(value)}" stroke="rgba(255,255,255,.1)"/><text x="${left - 8}" y="${y(value) + 4}" text-anchor="end" fill="#aaa6be" font-size="10">${compact.format(value)}</text>`;
    }
    svg += `<polyline points="${forecast.timeline.map((row, index) => `${x(index)},${y(row.revenue)}`).join(' ')}" fill="none" stroke="#6ee7b7" stroke-width="2.5" stroke-linejoin="round"/>`;
    svg += `<polyline points="${forecast.timeline.map((row, index) => `${x(index)},${y(row.adjustedRevenue)}`).join(' ')}" fill="none" stroke="#c4b5fd" stroke-width="2.5" stroke-dasharray="6 4" stroke-linejoin="round"/>`;
    forecast.timeline.forEach((row, index) => {
      svg += `<circle cx="${x(index)}" cy="${y(row.adjustedRevenue)}" r="1.8" fill="#c4b5fd"><title>${label(dates[index])}: ${money.format(row.adjustedRevenue)} after churn</title></circle>`;
      svg += `<circle cx="${x(index)}" cy="${y(row.revenue)}" r="2" fill="#6ee7b7"><title>${label(dates[index])}: ${money.format(row.revenue)} · ${row.sales} new sales</title></circle>`;
      if (index % 12 === 0 || index === 59) svg += `<text x="${x(index)}" y="${height - 17}" text-anchor="middle" fill="#aaa6be" font-size="10">${dates[index].toLocaleDateString('en-GB', { month: 'short', year: '2-digit', timeZone: 'UTC' })}</text>`;
    });
    get('chart').innerHTML = svg + '</svg>';
    get('months').innerHTML = forecast.timeline.map((row, index) => `<tr><td>${label(dates[index])}</td><td>${number.format(row.sales)}</td><td>${money.format(row.revenue)}</td><td>${money.format(row.adjustedRevenue)}</td></tr>`).join('');
  }
  function update() {
    const code = currency.value === 'USD' ? 'USD' : 'GBP';
    const money = new Intl.NumberFormat(code === 'USD' ? 'en-US' : 'en-GB', { style: 'currency', currency: code });
    document.querySelectorAll('[data-currency-symbol]').forEach(symbol => { symbol.textContent = code === 'USD' ? '$' : '£'; });
    set('currency-code', code);
    const selectedRate = get(scenario.value);
    const fields = ['leads', scenario.value, 'value', 'rate'].map(get);
    for (const key of Object.keys(scenarioNames)) get(key).dataset.selected = String(key === scenario.value);
    set('scenario-name', scenarioNames[scenario.value] + ' · ' + (selectedRate.value === '' ? '—' : selectedRate.value + '%'));
    set('top-period', label(dates[0]) + ' – ' + label(dates[11]) + ' | ' + code);
    set('value-description', mode.value === 'retainer' ? 'Monthly retainer revenue per client' : 'One-off revenue per deal');
    const applyChurn = churnEnabled.checked;
    get('churn-controls').hidden = !applyChurn;
    const values = fields.map(field => field.value === '' ? NaN : field.valueAsNumber);
    const valid = values.map((value, index) => Number.isFinite(value) && value >= 0 && (index !== 1 || value <= 100) && (index !== 3 || value > 0));
    fields.forEach((field, index) => field.setAttribute('aria-invalid', String(field.value !== '' && !valid[index])));
    const churnValue = applyChurn ? (churn.value === '' ? NaN : churn.valueAsNumber) : 0;
    const validChurn = Number.isFinite(churnValue) && churnValue >= 0 && churnValue <= 100;
    churn.setAttribute('aria-invalid', String(applyChurn && churn.value !== '' && !validChurn));
    get('result').dataset.state = '';
    set('error', '');
    pdf.disabled = true;
    exportSnapshot = null;
    function clear(message) {
      ['ratio', 'revenue', 'net', 'deals'].forEach(key => set(key, '—'));
      set('summary', message);
      set('chart', 'Complete your inputs to see the revenue forecast.');
      set('months', '');
      set('years', '');
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
    set('ratio', forecast.ratio.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' : 1');
    set('revenue', money.format(forecast.revenue));
    set('net', money.format(forecast.net));
    get('result').dataset.state = forecast.ratio < 1 ? 'below' : 'covered';
    set('summary', forecast.ratio < 1 ? 'Projected annual revenue is below your annual Unstuck investment.' : 'Projected annual revenue covers your annual Unstuck investment.');
    set('timing', `Your assumptions imply ${number.format(forecast.expectedDeals)} sales per year. This timeline books ${number.format(forecast.deals)} whole sales as they accumulate, starting with zero existing clients. ${applyChurn ? (mode.value === 'retainer' ? 'The separate churn scenario applies compound monthly retention to acquired retainer clients.' : 'The separate project churn scenario reduces future project revenue opportunities using compound monthly retention from the forecast start.') : 'Churn is off; both revenue lines are identical.'}`);
    get('years').innerHTML = forecast.years.map(row => `<tr><td>Year ${row.year}</td><td>${money.format(row.revenue)}</td><td>${money.format(row.accumulated)}</td><td>${money.format(row.adjustedRevenue)}</td><td>${row.ratio.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} : 1</td></tr>`).join('');
    drawChart(forecast, money, code);
    exportSnapshot = { forecast, inputs: { leads, conversion, value, rate, currency: code, scenario: scenarioNames[scenario.value], mode: mode.value, churn: churnValue, churnEnabled: applyChurn }, labels: dates.map(label) };
    pdf.disabled = false;
  }
  pdf.addEventListener('click', async () => {
    update();
    if (pdf.disabled || !exportSnapshot) return;
    const snapshot = exportSnapshot;
    pdf.disabled = true;
    pdf.textContent = 'PDF…';
    try {
      const source = get('pdf-logo');
      if (!source.complete) await new Promise((resolve, reject) => {
        source.addEventListener('load', resolve, { once: true });
        source.addEventListener('error', () => reject(new Error('The Unstuck logo could not load. Refresh and retry.')), { once: true });
      });
      if (!source.naturalWidth) throw new Error('The Unstuck logo could not load. Refresh and retry.');
      const canvas = document.createElement('canvas');
      canvas.width = source.naturalWidth; canvas.height = source.naturalHeight;
      canvas.getContext('2d').drawImage(source, 0, 0);
      const report = window.createRoiPdf({ ...snapshot, logo: canvas.toDataURL('image/png') }, window.jspdf.jsPDF);
      report.save(`Unstuck-Added-Revenue-${year}-${String(month + 1).padStart(2, '0')}.pdf`);
    } catch (error) {
      set('error', 'Could not create the PDF. ' + error.message);
    } finally {
      pdf.textContent = 'PDF ↓';
      pdf.disabled = !exportSnapshot;
    }
  });
  scenario.addEventListener('change', update);
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
    get('conservative').value = '3'; get('baseline').value = '5'; get('optimistic').value = '6';
    rates.GBP = '5490'; rates.USD = '7990';
    get('rate').value = rates[currency.value === 'USD' ? 'USD' : 'GBP'];
    update();
  }, 0));
  update();
})();
