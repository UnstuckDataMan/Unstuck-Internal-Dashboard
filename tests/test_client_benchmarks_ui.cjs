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
    { label: 'Mid-market', headcount: '51–500', average: 4.18637, display: '~5', sufficient: true, estimated: true },
    { label: 'Enterprise', headcount: '501+', average: 2.5562, display: '~3', sufficient: true, estimated: true },
  ],
  period: ['2026-04', '2026-05', '2026-06', '2026-07', '2026-08', '2026-09'],
  synced_at: '2026-10-27T10:00:00Z', polling_minutes: 30, stale: false,
});

async function mount(report = snapshot(), embedded = true) {
  const elements = {};
  for (const name of ['refresh', 'groups', 'note', 'updated', 'stale', 'status', 'message', 'report']) {
    elements['bench-' + name] = new Element();
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
  vm.runInNewContext(script, {
    document: { hidden: false, getElementById: id => elements[id], querySelector: () => view, createElement: tag => new Element(tag) },
    Event: UiEvent, Date, Number, String,
    AbortSignal: { timeout: () => ({}) },
    fetch: async url => { requests.push(url); return { ok: true, json: async () => current }; },
    setInterval: callback => intervals.push(callback),
  });
  await settle();
  assert.equal(elements['bench-report'].hidden, false, elements['bench-message'].textContent);
  return {
    elements, leads, selection, form, requests,
    buttons: () => descendants(elements['bench-groups'], 'button'),
    setReport: next => { current = next; },
    refresh: async () => { elements['bench-refresh'].dispatchEvent(new UiEvent('click')); await settle(); },
    poll: async () => { intervals[0](); await settle(); },
  };
}

async function main() {
  const ui = await mount();
  assert.equal(ui.buttons().length, 3);
  assert.equal(ui.leads.value, '5.00', 'Loading benchmarks must not change the forecast');
  assert.equal(ui.requests[0], '/api/expectations-setting/benchmarks');
  assert.match(ui.elements['bench-note'].textContent, /Rounded up for planning/);
  assert.match(ui.elements['bench-note'].textContent, /estimated activity and targeting inputs/);

  let forecastUpdates = 0;
  ui.form.addEventListener('input', () => { forecastUpdates++; });
  ui.buttons()[1].dispatchEvent(new UiEvent('click'));
  assert.equal(ui.leads.value, '4.19', 'Apply the source average to input precision, not the displayed ~5');
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
  assert.equal(ui.leads.value, '6.30', 'A subsequent deliberate apply uses the refreshed source');
  ui.form.dispatchEvent(new UiEvent('reset'));
  assert.equal(ui.selection.textContent, '', 'Reset must clear the selected benchmark context');

  const expired = snapshot();
  expired.stale = true;
  expired.groups.forEach(group => { group.display = 'Data too old'; });
  const old = await mount(expired);
  assert.ok(expired.groups.every(group => group.sufficient && Number.isFinite(group.average)));
  assert.equal(old.buttons().length, 0, 'Expired display gates must withhold applying still-present numeric averages');
  assert.equal(old.elements['bench-stale'].hidden, false);
  assert.equal(old.leads.value, '5.00');

  const limited = snapshot();
  limited.groups[1].sufficient = false;
  // Even an inconsistent cached display must not override the sample gate.
  const insufficient = await mount(limited);
  assert.equal(insufficient.buttons().length, 2);
  assert.equal(descendants(insufficient.elements['bench-groups'].children[1], 'button').length, 0);
  assert.equal((await mount(snapshot(), false)).buttons().length, 0, 'Standalone benchmarks must not expose calculator actions');
  console.log('Benchmark UI checks passed: freshness/sample gates, source-average application, bubbling updates, manual edits, refresh, reset and standalone isolation.');
}
main().catch(error => { console.error(error); process.exitCode = 1; });
