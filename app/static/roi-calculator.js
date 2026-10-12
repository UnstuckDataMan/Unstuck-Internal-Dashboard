(function () {
  'use strict';
  const form = document.getElementById('roi-form');
  const fields = ['leads', 'conversion', 'value', 'rate'].map(key => document.getElementById('roi-' + key));
  const currency = document.getElementById('roi-currency');
  const number = new Intl.NumberFormat('en-GB', { maximumFractionDigits: 3 });
  const set = (key, value) => { document.getElementById('roi-' + key).textContent = value; };
  function update() {
    const code = currency.value === 'USD' ? 'USD' : 'GBP';
    const money = new Intl.NumberFormat(code === 'USD' ? 'en-US' : 'en-GB', { style: 'currency', currency: code });
    document.querySelectorAll('[data-currency-symbol]').forEach(symbol => { symbol.textContent = code === 'USD' ? '$' : '£'; });
    set('currency-code', code);
    const values = fields.map(field => field.value === '' ? NaN : field.valueAsNumber);
    const valid = values.map((value, i) => Number.isFinite(value) && value >= 0 && (i !== 1 || value <= 100) && (i !== 3 || value > 0));
    fields.forEach((field, i) => field.setAttribute('aria-invalid', String(field.value !== '' && !valid[i])));
    const [leads, conversion, value, rate] = values;
    const deals = leads * conversion / 100;
    const annualDeals = deals * 12;
    set('deals', valid[0] && valid[1] && Number.isFinite(annualDeals) ? number.format(annualDeals) : '—');
    const result = document.getElementById('roi-result');
    result.dataset.state = '';
    set('error', '');
    if (!valid.every(Boolean)) {
      ['ratio', 'revenue', 'net'].forEach(key => set(key, '—'));
      const invalid = fields.findIndex((field, i) => field.value !== '' && !valid[i]);
      if (invalid !== -1) set('error', ['Enter a non-negative lead volume.', 'Enter a conversion rate between 0% and 100%.', 'Enter a non-negative conversion value.', 'Enter a monthly rate greater than zero.'][invalid]);
      set('summary', invalid !== -1 ? 'Check your figures to calculate the projection.' : 'Complete all four inputs to see your potential return.');
      return;
    }
    const revenue = deals * value;
    const annualRevenue = revenue * 12;
    const ratio = revenue / rate;
    if (![deals, annualDeals, revenue, annualRevenue, ratio, revenue - rate].every(Number.isFinite)) {
      ['deals', 'ratio', 'revenue', 'net'].forEach(key => set(key, '—'));
      set('error', 'These figures are too large. Enter smaller values.');
      set('summary', 'Check your figures to calculate the projection.');
      return;
    }
    set('ratio', ratio.toLocaleString('en-GB', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ':1');
    set('revenue', money.format(annualRevenue));
    set('net', money.format(revenue - rate));
    result.dataset.state = ratio < 1 ? 'below' : 'covered';
    set('summary', ratio < 1 ? 'Projected revenue is below your monthly investment.' : 'Projected revenue covers your monthly investment.');
  }
  currency.addEventListener('change', update);
  form.addEventListener('input', update);
  form.addEventListener('submit', event => event.preventDefault());
  form.addEventListener('reset', () => setTimeout(update, 0));
  update();
})();
