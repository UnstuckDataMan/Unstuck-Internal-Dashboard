const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const script = fs.readFileSync(path.join(__dirname, '../app/static/client-benchmarks.js'), 'utf8');
const settle = () => new Promise(resolve => setImmediate(resolve));

class UiEvent {
  constructor(type, options = {}) { this.type = type; this.bubbles = !!options.bubbles; }
}
class Element {
  constructor(tag = 'div') {
    this.tag = tag;
    this.children = [];
    this.handlers = {};
    this.attributes = {};
    this.dataset = {};
    this.textContent = '';
    this.value = '';
    this.classList = { add() {} };
  }
  append(...children) { children.forEach(child => { child.parentElement = this; this.children.push(child); }); }
  replaceChildren(...children) { this.children = []; this.append(...children); }
  setAttribute(name, value) { this.attributes[name] = String(value); }
  removeAttribute(name) { delete this.attributes[name]; }
  addEventListener(type, callback) { (this.handlers[type] ||= []).push(callback); }
  dispatchEvent(event) {
    (this.handlers[event.type] || []).forEach(callback => callback(event));
    if (event.bubbles && this.parentElement) this.parentElement.dispatchEvent(event);
    return true;
  }
  focus() { this.focused = true; }
}
const descendants = (element, tag) => element.children.flatMap(child => [
  ...(child.tag === tag ? [child] : []), ...descendants(child, tag),
]);
const snapshot = () => ({
  groups: [
    { label: 'Small businesses', headcount: '1–50', average: 8.6911, display: '~9', sufficient: true, estimated: true },
    { label: 'Mid-market', headcount: '51–1,000', average: 3.8013, display: '~4', sufficient: true, estimated: true },
    { label: 'Enterprise', headcount: '1,001+', average: 2.8254, display: '~3', sufficient: true, estimated: true },
  ],
  period: ['2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09'],
  synced_at: '2026-10-27T10:00:00Z', polling_minutes: 30, stale: false,
});

async function mount(report = snapshot(), embedded = true, newMarkup = true) {
  const elements = {};
  for (const name of ['refresh', 'groups', 'note', 'updated', 'stale', 'status', 'message', 'report']) {
    elements['bench-' + name] = new Element();
  }
  if (newMarkup) {
    for (const [name, text] of Object.entries({headline:'Average monthly leads.','intro-copy':'Recent client performance, grouped by company size.',metric:'Average leads',caption:'Average monthly leads by company size'})) {
      elements['bench-' + name] = new Element();
      elements['bench-' + name].textContent = text;
    }
    if (embedded) elements['bench-scale'] = new Element('button');
  }
  const leads = elements['roi-leads'] = new Element('input');
  const form = new Element('form');
  form.append(leads);
  leads.form = form;
  leads.value = '5.00';
  const selection = elements['roi-benchmark-selection'] = new Element();
  const view = new Element();
  view.dataset = { benchmarkEndpoint: '/api/expectations-setting/benchmarks', benchmarkCalculator: String(embedded) };
  const intervals = [];
  const requests = [];
  let current = report;
  let status = 200;
  vm.runInNewContext(script, {
    document: { hidden: false, getElementById: id => elements[id], querySelector: () => view, createElement: tag => new Element(tag) },
    Event: UiEvent, Date, Number, String,
    AbortSignal: { timeout: () => ({}) },
    fetch: async url => { requests.push(url); return { ok: status === 200, status, json: async () => current }; },
    setInterval: callback => intervals.push(callback),
  });
  await settle();
  assert.equal(elements['bench-report'].hidden, false, elements['bench-message'].textContent);
  return {
    elements, leads, selection, form, requests,
    buttons: () => descendants(elements['bench-groups'], 'button'),
    values: () => descendants(elements['bench-groups'], 'strong').map(element => element.textContent),
    toggle: () => elements['bench-scale'].dispatchEvent(new UiEvent('click')),
    setReport: next => { current = next; status = 200; },
    setFailure: (message, code = 503) => { current = {error: message}; status = code; },
    refresh: async () => { elements['bench-refresh'].dispatchEvent(new UiEvent('click')); await settle(); },
    poll: async () => { intervals[0](); await settle(); },
  };
}

async function main() {
  const ui = await mount();
  assert.equal(ui.buttons().length, 3);
  assert.equal(ui.leads.value, '5.00', 'Loading benchmarks must not change the forecast');
  assert.equal(ui.requests[0], '/api/expectations-setting/benchmarks');
  assert.doesNotMatch(ui.elements['bench-note'].textContent, /Rounded up for planning/);
  assert.match(ui.elements['bench-note'].textContent, /estimated activity and targeting inputs/);

  let forecastUpdates = 0;
  ui.form.addEventListener('input', () => { forecastUpdates++; });
  ui.buttons()[1].dispatchEvent(new UiEvent('click'));
  assert.equal(ui.leads.value, '4', 'Apply the displayed whole-lead planning figure, not the underlying average');
  assert.equal(forecastUpdates, 1, 'The input event must bubble so the existing forecast recalculates');
  assert.equal(ui.selection.textContent, 'Based on Mid-market · Apr 2026 — Sept 2026');
  assert.equal(ui.leads.focused, true);

  ui.leads.value = '6.75';
  ui.leads.dispatchEvent(new UiEvent('input', { bubbles: true }));
  assert.equal(ui.selection.textContent, '', 'A manual assumption must not retain benchmark attribution');
  const newer = snapshot();
  newer.groups[1].average = 6.3;
  newer.groups[1].display = '~7';
  ui.setReport(newer);
  await ui.refresh();
  await ui.poll();
  assert.equal(ui.leads.value, '6.75', 'Manual and background refreshes must preserve the entered assumption');
  assert.equal(ui.selection.textContent, '');
  assert.equal(forecastUpdates, 2, 'Refreshing data must not dispatch a forecast update');
  ui.buttons()[1].dispatchEvent(new UiEvent('click'));
  assert.equal(ui.leads.value, '7', 'A subsequent deliberate apply uses the refreshed rounded display');
  ui.form.dispatchEvent(new UiEvent('reset'));
  assert.equal(ui.selection.textContent, '', 'Reset must clear the selected benchmark context');

  const expired = snapshot();
  expired.stale = true;
  expired.groups.forEach(group => { group.display = 'Data too old'; });
  const old = await mount(expired);
  assert.ok(expired.groups.every(group => group.sufficient && Number.isFinite(group.average)));
  assert.equal(old.buttons().length, 0, 'Expired display gates must withhold applying still-present numeric averages');
  assert.equal(old.elements['bench-stale'].hidden, false);
  assert.equal(old.elements['bench-scale'].disabled, true);
  assert.equal(old.leads.value, '5.00');

  const limited = snapshot();
  limited.groups[1].sufficient = false;
  // Even an inconsistent cached display must not override the sample gate.
  const insufficient = await mount(limited);
  assert.equal(insufficient.buttons().length, 2);
  assert.equal(descendants(insufficient.elements['bench-groups'].children[1], 'button').length, 0);
  const standalone = await mount(snapshot(), false);
  assert.equal(standalone.buttons().length, 0, 'Standalone benchmarks must not expose calculator actions');
  assert.equal(standalone.elements['bench-scale'], undefined);
  assert.deepEqual(standalone.values(), ['~9','~4','~3']);
  const rounded = await mount();
  for (const [index, expected] of ['9', '4', '3'].entries()) {
    rounded.buttons()[index].dispatchEvent(new UiEvent('click'));
    assert.equal(rounded.leads.value, expected, 'Every tier must apply its displayed planning figure');
  }
  const cachedTemplate = await mount(snapshot(), true, false);
  assert.equal(cachedTemplate.buttons().length, 3, 'Optional new markup must not break older cached templates');
  assert.deepEqual(cachedTemplate.values(), ['~9','~4','~3']);

  const source = snapshot(), unchanged = JSON.stringify(source);
  const scaled = await mount(source);
  const scaleButton = scaled.elements['bench-scale'];
  assert.equal(scaleButton.disabled, false);
  assert.equal(scaleButton.attributes['aria-pressed'], 'false');
  const requestsBefore = scaled.requests.length;
  scaled.toggle();
  assert.deepEqual(scaled.values(), ['~27','~12','~9']);
  assert.equal(scaleButton.attributes['aria-pressed'], 'true');
  assert.equal(scaled.elements['bench-headline'].textContent, 'Scaled monthly leads.');
  assert.equal(scaled.elements['bench-metric'].textContent, 'Scaled leads');
  assert.equal(scaled.elements['bench-intro-copy'].textContent, 'Explore a 3× lead-volume scenario for your campaign.');
  assert.match(scaled.elements['bench-caption'].textContent, /Scaled scenario/);
  assert.equal(scaled.buttons()[1].attributes['aria-label'], 'Use Mid-market scaled scenario in forecast');
  assert.match(scaled.elements['bench-note'].textContent, /^Scaled scenario: 3× benchmark volumes\. Six-month benchmarks using estimated/);
  assert.equal(scaled.leads.value, '5.00', 'Scale toggles must not silently change the calculator');
  for (let i=0;i<10;i++) {
    scaled.toggle();assert.deepEqual(scaled.values(), ['~9','~4','~3']);
    scaled.toggle();assert.deepEqual(scaled.values(), ['~27','~12','~9']);
  }
  assert.equal(scaled.requests.length, requestsBefore, 'Toggling must render the cached source without fetching');
  for (const [index, expected] of ['27','12','9'].entries()) {
    scaled.buttons()[index].dispatchEvent(new UiEvent('click'));
    assert.equal(scaled.leads.value, expected, 'Apply the visible scaled rounded volume, not three times the source average');
    assert.match(scaled.selection.textContent, /Scale client 3×/);
  }
  assert.equal(JSON.stringify(source), unchanged, 'Scale presentation must not mutate any source data');
  scaled.leads.value='6.75';scaled.leads.dispatchEvent(new UiEvent('input',{bubbles:true}));
  const refreshData=snapshot();refreshData.groups[1].average=6.3;refreshData.groups[1].display='~7';
  scaled.setReport(refreshData);await scaled.refresh();await scaled.poll();
  assert.deepEqual(scaled.values(), ['~27','~21','~9']);
  assert.equal(scaleButton.attributes['aria-pressed'], 'true', 'Fresh background updates preserve the selected scenario');
  assert.equal(scaled.leads.value, '6.75');
  assert.equal(scaled.selection.textContent, '');
  scaled.form.dispatchEvent(new UiEvent('reset'));
  assert.deepEqual(scaled.values(), ['~9','~7','~3']);
  assert.equal(scaleButton.attributes['aria-pressed'], 'false');
  assert.equal(scaled.elements['bench-headline'].textContent, 'Average monthly leads.');
  assert.equal(scaled.elements['bench-metric'].textContent, 'Average leads');
  assert.equal(scaled.elements['bench-intro-copy'].textContent, 'Recent client performance, grouped by company size.');
  assert.doesNotMatch(scaled.elements['bench-note'].textContent, /Scaled scenario/);
  assert.equal(scaled.selection.textContent, '');

  scaled.toggle();scaled.setReport(expired);await scaled.refresh();
  assert.equal(scaleButton.disabled,true);assert.equal(scaleButton.attributes['aria-pressed'],'false');
  assert.deepEqual(scaled.values(), ['Data too old','Data too old','Data too old']);
  assert.equal(scaled.buttons().length,0);scaled.toggle();assert.equal(scaleButton.attributes['aria-pressed'],'false');
  const stale=snapshot();stale.stale=true;
  scaled.setReport(stale);await scaled.refresh();
  assert.equal(scaleButton.disabled,true,'A stale snapshot must not enable Scale Client');
  assert.deepEqual(scaled.values(), ['~9','~4','~3']);
  const noSamples=snapshot();noSamples.groups.forEach(group=>{group.sufficient=false;});
  scaled.setReport(noSamples);await scaled.refresh();assert.equal(scaleButton.disabled,true);assert.equal(scaled.buttons().length,0);
  scaled.setReport(snapshot());await scaled.refresh();scaled.toggle();
  scaled.setFailure('Sheet unavailable');await scaled.refresh();
  assert.equal(scaleButton.disabled,true);assert.equal(scaleButton.attributes['aria-pressed'],'false');
  assert.equal(scaled.elements['bench-report'].hidden,true);
  assert.equal(scaled.elements['bench-message'].textContent,'Sheet unavailable');
  scaled.toggle();scaled.form.dispatchEvent(new UiEvent('reset'));
  assert.equal(scaled.elements['bench-report'].hidden,true,'Reset must not resurrect a report after a failed fetch');
  assert.equal(scaled.leads.value,'6.75');
  scaled.setReport(snapshot());await scaled.refresh();
  assert.equal(scaleButton.disabled,false);assert.deepEqual(scaled.values(), ['~9','~4','~3']);
  console.log('Benchmark UI checks passed: freshness/sample gates, rounded planning figures, bubbling updates, manual edits, refresh, reset and standalone isolation.');
  console.log('Scale Client checks passed: reversible 3× rounded volumes, immutable source, exact forecast application, refresh/manual-input isolation, reset, expiry and error recovery.');
}
main().catch(error => { console.error(error); process.exitCode = 1; });
