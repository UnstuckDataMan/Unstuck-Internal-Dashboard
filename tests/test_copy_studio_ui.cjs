const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.join(__dirname, '..');
const settle = async () => { for (let i = 0; i < 5; i++) await new Promise(resolve => setImmediate(resolve)); };
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => { resolve = yes; reject = no; });
  return {promise, resolve, reject};
};

class UiEvent {
  constructor(type, options = {}) { Object.assign(this, {type, bubbles: false, defaultPrevented: false}, options); }
  preventDefault() { this.defaultPrevented = true; }
  stopPropagation() { this.bubbles = false; }
}

const descendants = element => element.children.flatMap(child => [child, ...descendants(child)]);
const dataKey = value => value.replace(/-([a-z])/g, (_, letter) => letter.toUpperCase());
function matches(element, selector) {
  if (selector.includes(',')) return selector.split(',').some(part => matches(element, part.trim()));
  selector = selector.trim();
  const last = selector.split(/\s+/).at(-1);
  if (selector !== last) return matches(element, last);
  if (/:checked\b/.test(selector) && !element.checked) return false;
  if (/:disabled\b/.test(selector) && !element.disabled) return false;
  selector = selector.replace(/:(checked|disabled)\b/g, '');
  const simple = selector.replace(/\[[^\]]*\]/g, '');
  const tag = simple.match(/^[\w-]+/);
  if (tag && element.tagName !== tag[0].toUpperCase()) return false;
  const id = simple.match(/#([\w-]+)/);
  if (id && element.id !== id[1]) return false;
  for (const cls of simple.matchAll(/\.([\w-]+)/g)) if (!element.classList.contains(cls[1])) return false;
  for (const attr of selector.matchAll(/\[([^\]=\s]+)(?:=["']?([^\]"']*)["']?)?\]/g)) {
    const value = element.getAttribute(attr[1]);
    if (value === null || (attr[2] !== undefined && value !== attr[2])) return false;
  }
  return true;
}

class Element {
  constructor(tag = 'div') {
    this.tagName = tag.toUpperCase();
    this.children = [];
    this.parentElement = null;
    this.handlers = {};
    this.attributes = {};
    this.dataset = {};
    this.style = {};
    this.value = '';
    this.checked = false;
    this.disabled = false;
    this.hidden = false;
    this._text = '';
    this.className = '';
    this.classList = {
      contains: value => this.className.split(/\s+/).includes(value),
      add: (...values) => { this.className = [...new Set([...this.className.split(/\s+/).filter(Boolean), ...values])].join(' '); },
      remove: (...values) => { this.className = this.className.split(/\s+/).filter(value => !values.includes(value)).join(' '); },
      toggle: (value, force) => {
        const active = force === undefined ? !this.classList.contains(value) : force;
        this.classList[active ? 'add' : 'remove'](value);
        return active;
      },
    };
  }
  get textContent() { return this._text + this.children.map(child => child.textContent).join(''); }
  set textContent(value) { this._text = String(value ?? ''); this.children = []; }
  set innerHTML(_) { throw new Error('Unsafe innerHTML rendering is forbidden in Copy Studio'); }
  append(...children) {
    for (let child of children) {
      if (typeof child === 'string') { const text = new Element('#text'); text.textContent = child; child = text; }
      child.parentElement = this;
      this.children.push(child);
    }
  }
  appendChild(child) { this.append(child); return child; }
  replaceChildren(...children) { this.children = []; this._text = ''; this.append(...children); }
  remove() { if (this.parentElement) this.parentElement.children = this.parentElement.children.filter(child => child !== this); }
  setAttribute(name, value) {
    value = String(value);
    this.attributes[name] = value;
    if (name.startsWith('data-')) this.dataset[dataKey(name.slice(5))] = value;
    else if (name === 'class') this.className = value;
    else if (['disabled', 'hidden', 'checked', 'required', 'selected'].includes(name)) this[name] = true;
    else this[name] = value;
  }
  getAttribute(name) {
    if (name.startsWith('data-')) return this.dataset[dataKey(name.slice(5))] ?? null;
    if (name === 'class') return this.className || null;
    if (['disabled', 'hidden', 'checked', 'required', 'selected'].includes(name)) return this[name] ? '' : null;
    return this.attributes[name] ?? this[name] ?? null;
  }
  removeAttribute(name) { delete this.attributes[name]; if (name in this) this[name] = false; }
  addEventListener(type, callback) { (this.handlers[type] ||= []).push(callback); }
  dispatchEvent(event) {
    event.target ||= this;
    event.currentTarget = this;
    for (const callback of this.handlers[event.type] || []) callback(event);
    if (event.bubbles && this.parentElement) this.parentElement.dispatchEvent(event);
    return !event.defaultPrevented;
  }
  click() { if (!this.disabled) this.dispatchEvent(new UiEvent('click', {bubbles: true})); }
  focus() { this.focused = true; }
  scrollIntoView() {}
  querySelectorAll(selector) { return descendants(this).filter(element => matches(element, selector)); }
  querySelector(selector) { return this.querySelectorAll(selector)[0] || null; }
  closest(selector) { return matches(this, selector) ? this : this.parentElement?.closest(selector) || null; }
  get options() { return this.children.filter(child => child.tagName === 'OPTION'); }
  get selectedOptions() { return this.options.filter(option => option.value === this.value); }
  get valueAsNumber() { return this.value === '' ? NaN : Number(this.value); }
  checkValidity() { return !this.required || !!this.value.trim(); }
  reportValidity() { return this.checkValidity(); }
}

function documentFromTemplate(html) {
  const body = new Element('body'), stack = [body];
  const voidTags = new Set(['INPUT', 'BR', 'HR', 'IMG', 'LINK', 'META']);
  for (const token of html.matchAll(/<\/?([a-z][\w-]*)\b([^>]*)>|([^<]+)/gi)) {
    if (!token[1]) {
      if (token[3].trim()) stack.at(-1).append(token[3]);
      continue;
    }
    if (token[0].startsWith('</')) {
      const index = stack.findLastIndex(element => element.tagName === token[1].toUpperCase());
      if (index > 0) stack.length = index;
      continue;
    }
    const element = new Element(token[1]);
    for (const attr of token[2].matchAll(/([^\s=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g)) {
      element.setAttribute(attr[1], attr[2] ?? attr[3] ?? attr[4] ?? '');
    }
    stack.at(-1).append(element);
    if (!voidTags.has(element.tagName) && !token[0].endsWith('/>')) stack.push(element);
  }
  for (const element of descendants(body)) {
    if (element.tagName === 'SELECT') element.value = (element.options.find(option => option.selected) || element.options[0])?.value || '';
    if (element.tagName === 'TEXTAREA') element.value = element.textContent.trim();
    element.form = element.closest('form');
  }
  return {
    body, hidden: false,
    getElementById: id => descendants(body).find(element => element.id === id) || null,
    querySelector: selector => body.querySelector(selector),
    querySelectorAll: selector => body.querySelectorAll(selector),
    createElement: tag => new Element(tag),
    createTextNode: text => { const element = new Element('#text'); element.textContent = text; return element; },
    createDocumentFragment: () => new Element('fragment'),
  };
}

const response = (data, status = 200) => ({ok: status >= 200 && status < 300, status, json: async () => data});

// The scenarios below exercise the actual script through form and button events.
// Responses are synthetic fixtures: no application or provider network is used.
const malicious = '<img src=x onerror="globalThis.injected=true">';
function record(id, kind, content, overrides = {}) {
  return {id, kind, client_id: 'client-a', source_key: `synthetic:${id}`, content_hash: `hash-${id}`, parent_id: null,
    created_by: 'writer@example.invalid', created_at: '2026-10-28T12:00:00Z', status: 'approved', content, ...overrides};
}
function clientData() {
  return {
    brief: record('brief-1', 'brief', {
      facts: [{id: 'fact-1', category: 'offer', key: 'Core offer', text: `Client evidence ${malicious}`,
        evidence: [{source_id: 'source-1', block: '1', quote: 'Verified fixture quote', start_ms: 0, end_ms: 1000}]}],
      constraints: ['Keep claims specific and sourced.'], contradictions: [], missing_fields: [],
      coverage: {complete: true, total_chunks: 1, processed_chunks: 1, sources: ['source-1']},
    }),
    sources: [record('source-1', 'source', {title: malicious, text: 'Verified fixture quote', url: 'javascript:alert(1)'})],
    examples: [record('example-1', 'example', {title: 'Approved example', channel: 'email', source: 'copy_bank', sanitized: true, text: 'A sourced example.'}),
      record('example-li', 'example', {title: 'Approved LinkedIn example', channel: 'linkedin', source: 'copy_bank', sanitized: true, text: 'A sourced LinkedIn example.'})],
    drafts: [record('draft-1', 'draft', {
      channel: 'email', mode: 'draft',
      angles: [{name: `Specific relevance ${malicious}`, rationale: 'Connect the verified offer to the selected audience.', subjects: ['A relevant question'],
        steps: [{kind: 'initial', body: 'Hello {{first_name}}, a relevant conversation?', evidence_ids: ['fact-1']},
          {kind: 'followup', body: 'Following up with one useful detail.', evidence_ids: ['fact-1']}], evidence_ids: ['fact-1']}],
      review: {issues: [], checks: [], can_approve: true}, provenance: {brief_id: 'brief-1', example_ids: ['example-1'], source_ids: ['source-1'],
        limits: {subject: 120, initial: 1200, follow_up: 800, closing: 500}},
    }, {status: 'draft'})],
    pending_briefs: [], methodology: [], performance: [], jobs: [], completeness: {complete: true, score: 100, missing_fields: [], total_chunks: 1, processed_chunks: 1},
  };
}

function zeroCosts() {
  const bucket = () => ({amount:0,known_cost:0,total_calls:0,known_calls:0,unpriced_calls:0,unknown_usage_calls:0,complete:true});
  return {currency:'USD',timezone:'Europe/London',today:bucket(),month:bucket(),all_time:bucket()};
}
async function mount({canReview = true, ready = true, modelConfigured = true, data = clientData(), costs = zeroCosts()} = {}) {
  const document = documentFromTemplate(fs.readFileSync(path.join(root, 'app/templates/copy_studio.html'), 'utf8'));
  const requests = [], timers = new Map(), routes = new Map(), clipboard = [];
  let nextTimer = 1;
  const bootstrap = {clients: [{id: 'client-a', name: `Alpha ${malicious}`}, {id: 'client-b', name: 'Beta'}],
    capabilities: [], ready, can_review: canReview, csrf: 'csrf-fixture-token', model: {configured: modelConfigured, provider: 'openai'},costs};
  routes.set('GET /api/copy-studio/bootstrap', () => response(bootstrap));
  routes.set('GET /api/copy-studio/costs', () => response(costs));
  routes.set('GET /api/copy-studio/clients/client-a', () => response(data));
  routes.set('GET /api/copy-studio/clients/client-b', () => response({...clientData(), brief: null, drafts: [], sources: [], examples: []}));
  const fetch = async (url, options = {}) => {
    const method = options.method || 'GET';
    const request = {url: String(url), method, body: options.body ? JSON.parse(options.body) : undefined, signal: options.signal,
      headers: options.headers, credentials: options.credentials};
    requests.push(request);
    if (method === 'POST') {
      assert.equal(options.headers?.['X-Copy-Studio-CSRF'], bootstrap.csrf, 'Every mutation must carry the bootstrap CSRF token');
      assert.equal(options.credentials, 'same-origin', 'Mutation credentials must stay within the dashboard origin');
    }
    const handler = routes.get(`${method} ${url}`);
    if (!handler) throw new Error(`Unexpected offline request: ${method} ${url}`);
    return handler(request);
  };
  const schedule = (callback, delay) => { const id = nextTimer++; timers.set(id, {callback, delay}); return id; };
  const window = {location: {href: 'https://dashboard.example.invalid/copy-studio'}, addEventListener() {}, confirm: () => true};
  const context = {document, window, fetch, Event: UiEvent, AbortController, AbortSignal, URL,
    Intl, Date, Number, String, Object, Array, JSON, Math, console,
    navigator: {clipboard: {writeText: async text => { clipboard.push(text); }}},
    setTimeout: schedule, clearTimeout: id => timers.delete(id), setInterval: schedule, clearInterval: id => timers.delete(id),
    confirm: () => true, structuredClone,
  };
  vm.runInNewContext(fs.readFileSync(path.join(root, 'app/static/copy-studio.js'), 'utf8'), context);
  await settle();
  const get = id => document.getElementById(id.startsWith('cs-') ? id : `cs-${id}`);
  return {
    get, document, requests, timers, routes, context, clipboard, data,
    route: (method, url, handler) => routes.set(`${method} ${url}`, handler),
    async chooseClient(id) { get('client').value = id; get('client').dispatchEvent(new UiEvent('change')); await settle(); },
    async change(id, value, type = 'input') { get(id).value = value; get(id).dispatchEvent(new UiEvent(type, {bubbles: true})); await settle(); },
    async click(id) { get(id).click(); await settle(); },
    async submit(id) { get(id).dispatchEvent(new UiEvent('submit', {bubbles: true})); await settle(); },
    async tick() {
      const entry = timers.entries().next().value;
      if (!entry) throw new Error('Expected a scheduled job poll');
      timers.delete(entry[0]); entry[1].callback(); await settle();
    },
  };
}

function draftReviewButton(ui, decision) {
  return ui.get('draft-review').querySelector(`[data-decision="${decision}"]`);
}

async function testSafeRenderingAndClientIsolation() {
  const ui = await mount();
  await ui.chooseClient('client-a');
  assert.equal(ui.get('workspace').hidden, false);
  assert.ok(ui.document.body.textContent.includes(malicious), 'Untrusted client, fact and draft strings must appear as literal text');
  assert.equal(ui.document.querySelectorAll('img').length, 0, 'Source strings must never become executable markup');
  assert.ok(ui.document.querySelectorAll('a').every(link => !/^javascript:/i.test(link.href || link.getAttribute('href') || '')));
  assert.equal(ui.context.injected, undefined);

  const pending = deferred();
  ui.route('GET', '/api/copy-studio/clients/client-a', () => pending.promise);
  await ui.chooseClient('client-a');
  await ui.chooseClient('client-b');
  pending.resolve(response(clientData()));
  await settle();
  assert.equal(ui.get('client').value, 'client-b');
  assert.ok(!ui.get('facts').textContent.includes('Client evidence'), 'A late response must not expose the previous client’s brief');
  assert.equal(ui.get('angles').querySelectorAll('textarea').length, 0, 'A late response must not replace another client’s draft');
  assert.equal(ui.get('generate').disabled, true, 'A client without an approved brief must not inherit generation readiness');
}

async function testGenerationAndPollingStayClientScoped() {
  const ui = await mount();
  await ui.chooseClient('client-a');
  ui.route('POST', '/api/copy-studio/clients/client-a/generate', () => response({id: 'job-1', client_id: 'client-a', status: 'running', kind: 'generate'}));
  await ui.change('channel', 'email', 'change');
  await ui.change('mode', 'assemble', 'change');
  for (const [id, value] of Object.entries({'subsector': 'B2B software', geography: 'United Kingdom', 'company-size': '51–200', role: 'Marketing Director', offer: 'A sourced campaign offer'})) await ui.change(id, value);
  const example = ui.get('examples').querySelector('input[type="checkbox"]');
  assert.ok(example, 'An approved example should be selectable');
  example.checked = true; example.dispatchEvent(new UiEvent('change', {bubbles: true}));
  await ui.submit('generate-form');
  const generated = ui.requests.find(request => request.method === 'POST' && request.url.endsWith('/generate'));
  assert.ok(generated, ui.get('status').textContent);
  assert.equal(generated.body.channel, 'email');
  assert.equal(generated.body.mode, 'assemble');
  assert.deepEqual(generated.body.audience, {subsector: 'B2B software', geography: 'United Kingdom', company_size: '51–200', role: 'Marketing Director', offer: 'A sourced campaign offer'});
  assert.deepEqual(generated.body.example_ids, ['example-1']);
  assert.deepEqual(generated.body.limits, {subject: 120, initial: 1200, follow_up: 800, closing: 500});
  assert.equal(Object.hasOwn(generated.body, 'provider'), false, 'The provider is server-configured, not a client-selected generation field');
  assert.ok(ui.timers.size, 'A running generation job should schedule progress polling');
  await ui.chooseClient('client-b');
  const before = ui.requests.length;
  if (ui.timers.size) await ui.tick();
  assert.equal(ui.requests.slice(before).filter(request => request.url.includes('/clients/client-a/')).length, 0,
    'Switching clients must cancel or ignore previous client job polls');
  assert.equal(ui.get('client').value, 'client-b');

  const linkedin = await mount();
  await linkedin.chooseClient('client-a');
  linkedin.route('POST', '/api/copy-studio/clients/client-a/generate', () => response({id: 'job-li', client_id: 'client-a', status: 'running', kind: 'generate'}));
  await linkedin.change('channel', 'linkedin', 'change');
  await linkedin.change('mode', 'draft', 'change');
  await linkedin.change('role', 'Operations Director');
  await linkedin.change('offer', 'The approved client offer');
  await linkedin.submit('generate-form');
  const linkedinRequest = linkedin.requests.find(request => request.method === 'POST' && request.url.endsWith('/generate'));
  assert.ok(linkedinRequest, linkedin.get('status').textContent);
  assert.equal(linkedinRequest.body.channel, 'linkedin');
  assert.equal(linkedinRequest.body.mode, 'draft');
  assert.deepEqual(linkedinRequest.body.limits, {connection: 200, initial: 600, follow_up: 600, closing: 400});
  assert.deepEqual(linkedinRequest.body.example_ids, []);
}

async function testImmutableRevisionsAndApprovalGate() {
  const data = clientData(), initialContent = JSON.stringify(data.drafts[0].content);
  const ui = await mount({data});
  await ui.chooseClient('client-a');
  const body = ui.get('angles').querySelector('[data-copy-path="angles.0.steps.0.body"]');
  assert.ok(body, 'Draft messages must be editable');
  const approval = draftReviewButton(ui, 'approved');
  assert.ok(approval && !approval.disabled, 'A reviewer can approve a valid saved draft');
  body.value = 'Updated wording grounded in the verified offer.';
  body.dispatchEvent(new UiEvent('input', {bubbles: true}));
  await settle();
  assert.equal(draftReviewButton(ui, 'approved').disabled, true, 'Unsaved copy cannot approve the previous immutable version');
  assert.equal(JSON.stringify(data.drafts[0].content), initialContent, 'Editing must not mutate the loaded source version');
  assert.equal(ui.get('save-draft').disabled, false);
  ui.route('POST', '/api/copy-studio/clients/client-a/drafts/draft-1/revisions', request => {
    const saved = record('draft-2', 'draft', request.body.content, {status: 'draft', parent_id: 'draft-1'});
    data.drafts.unshift(saved);
    return response(saved, 201);
  });
  await ui.click('save-draft');
  const revision = ui.requests.find(request => request.url.endsWith('/drafts/draft-1/revisions'));
  assert.ok(revision, ui.get('status').textContent);
  assert.equal(revision.method, 'POST');
  assert.equal(revision.body.content.angles[0].steps[0].body, 'Updated wording grounded in the verified offer.');
  assert.equal(data.drafts[0].parent_id, 'draft-1');
  assert.equal(JSON.stringify(data.drafts[1].content), initialContent, 'Saving a revision must preserve the original draft');

  ui.route('POST', '/api/copy-studio/clients/client-a/records/draft-2/review', request => {
    data.drafts[0].status = request.body.decision;
    return response(data.drafts[0]);
  });
  const notes = ui.get('draft-review').querySelector('textarea');
  if (notes) notes.value = 'Reviewed against the cited client facts.';
  draftReviewButton(ui, 'approved').click();
  await settle();
  const reviewed = ui.requests.find(request => request.url.endsWith('/records/draft-2/review'));
  assert.ok(reviewed, 'Approval must target the saved revision, not its parent');
  assert.equal(reviewed.body.decision, 'approved');
  assert.equal(typeof reviewed.body.notes, 'string');

  const overlengthData = clientData();
  overlengthData.drafts[0].content.angles[0].steps[0].body = 'x'.repeat(1201);
  const overlength = await mount({data: overlengthData});
  await overlength.chooseClient('client-a');
  assert.equal(draftReviewButton(overlength, 'approved').disabled, true,
    'An already saved draft over the channel limit must not be approved even if server review says it is valid');

  const viewer = await mount({canReview: false});
  await viewer.chooseClient('client-a');
  const viewerApproval = draftReviewButton(viewer, 'approved');
  assert.ok(!viewerApproval || viewerApproval.disabled, 'Writer-only users must not receive an enabled approval action');
  const blockedData = clientData();
  blockedData.drafts[0].content.review = {can_approve: false, issues: [{severity: 'error', message: 'Unsupported claim'}], checks: []};
  const blocked = await mount({data: blockedData});
  await blocked.chooseClient('client-a');
  assert.equal(draftReviewButton(blocked, 'approved').disabled, true, 'Server review failures must keep approval disabled');
  assert.match(blocked.get('review-issues').textContent, /Unsupported claim/);
}

async function testCompletedJobLoadsItsSavedVersion() {
  const data = clientData(), ui = await mount({data});
  await ui.chooseClient('client-a');
  ui.route('POST', '/api/copy-studio/clients/client-a/generate', () => response({id: 'job-complete', client_id: 'client-a', status: 'running'}));
  ui.route('GET', '/api/copy-studio/clients/client-a/jobs/job-complete', () => {
    const content = structuredClone(data.drafts[0].content);
    content.angles[0].name = 'A newly generated distinct angle';
    data.drafts.unshift(record('generated-2', 'draft', content, {status: 'draft'}));
    return response({id: 'job-complete', client_id: 'client-a', status: 'completed', result: {record_id: 'generated-2'}});
  });
  await ui.change('role', 'Operations Director');
  await ui.change('offer', 'The approved client offer');
  await ui.submit('generate-form');
  await ui.tick();
  assert.equal(ui.get('version').value, 'generated-2', 'A completed job must open its actual saved result');
  assert.match(ui.get('angles').textContent, /A newly generated distinct angle/);
  assert.ok(ui.requests.some(request => request.url.endsWith('/jobs/job-complete')));
  assert.equal(ui.timers.size, 0, 'A terminal job must stop polling');
}

async function testPendingBriefEditingAndReviewAreVersionSpecific() {
  const data = clientData(), pending = structuredClone(data.brief);
  pending.id = 'brief-pending'; pending.status = 'pending'; pending.created_at = '2026-10-28T13:00:00Z';
  pending.content.facts[0].text = 'A newly extracted fact awaiting review.';
  data.pending_briefs = [pending];
  const ui = await mount({data});
  await ui.chooseClient('client-a');
  assert.equal(ui.get('brief-version').value, pending.id, 'The newest pending brief must be available for review');
  assert.equal(ui.get('facts').querySelector('textarea').value, pending.content.facts[0].text);
  assert.equal(ui.get('generate').disabled, false, 'A pending review must not replace the still-approved generation brief');
  await ui.change('brief-version', 'brief-1', 'change');
  assert.equal(ui.get('facts').querySelector('textarea').value, data.brief.content.facts[0].text);
  await ui.change('brief-version', pending.id, 'change');
  ui.route('POST', '/api/copy-studio/clients/client-a/records', request => {
    const saved = record('brief-revised', 'brief', request.body.content, {status: 'pending', parent_id: request.body.parent_id, created_at: '2026-10-28T14:00:00Z'});
    data.pending_briefs.unshift(saved);
    return response(saved, 201);
  });
  await ui.change('brief-constraints', 'Keep the newly verified offer specific.');
  await ui.submit('brief-form');
  const saved = ui.requests.find(request => request.method === 'POST' && request.url.endsWith('/records'));
  assert.equal(saved.body.parent_id, 'brief-pending');
  assert.equal(saved.body.content.facts[0].text, pending.content.facts[0].text);
  assert.equal(data.brief.id, 'brief-1', 'Saving a pending version must not change the approved brief');
  assert.equal(ui.get('brief-version').value, 'brief-revised');
  ui.route('POST', '/api/copy-studio/clients/client-a/records/brief-revised/review', request => {
    const approved = data.pending_briefs.find(item => item.id === 'brief-revised');
    approved.status = request.body.decision; data.brief = approved; data.pending_briefs = [];
    return response(approved);
  });
  ui.get('brief-review').querySelector('[data-decision="approved"]').click();
  await settle();
  assert.ok(ui.requests.some(request => request.url.endsWith('/records/brief-revised/review')));
  assert.equal(ui.get('brief-status').textContent, 'approved');
}

async function testTranscriptInspectionAndCompletenessConfirmation() {
  const data = clientData();
  const transcript = record('transcript-1', 'transcript', {title: 'Call transcript', source: 'read_ai', source_complete: false, text: 'Summary only'}, {status: 'pending'});
  data.sources = [transcript];
  data.examples.push(record('example-unsafe', 'example', {channel: 'email', subject: 'Private recipient', body: 'Recipient-specific content', sanitized: false}, {status: 'pending'}));
  const ui = await mount({data});
  await ui.chooseClient('client-a');
  const card = ui.document.getElementById('cs-source-transcript-1');
  assert.equal(card.querySelectorAll('[data-decision]').length, 0, 'Raw transcripts must not receive generic record approval actions');
  assert.equal(ui.document.getElementById('cs-source-example-unsafe').querySelector('[data-decision="approved"]').disabled, true,
    'Unsanitized examples must not be directly approved for reuse');
  ui.route('GET', '/api/copy-studio/clients/client-a/records/transcript-1', () => response({record: {...transcript, content: {...transcript.content, transcript: [{speaker: 'Client', text: `Full transcript ${malicious}`} ]}}}));
  card.querySelector('[data-view-record="transcript-1"]').click();
  await settle();
  assert.match(card.textContent, /Full transcript/);
  assert.ok(card.textContent.includes(malicious));
  assert.equal(card.querySelectorAll('img').length, 0);
  const check = card.querySelector('[data-confirm-source="transcript-1"]');
  const confirm = card.querySelector('[data-confirm-record="transcript-1"]');
  assert.equal(confirm.disabled, true);
  check.checked = true; check.dispatchEvent(new UiEvent('change'));
  await ui.change('confirm-notes-transcript-1', 'Checked');
  assert.equal(confirm.disabled, true, 'A checkbox without meaningful completeness notes is insufficient');
  const notes = 'Compared all transcript blocks with the original recorded call.';
  await ui.change('confirm-notes-transcript-1', notes);
  assert.equal(confirm.disabled, false);
  ui.route('POST', '/api/copy-studio/clients/client-a/transcripts/transcript-1/confirm-complete', () => {
    transcript.content.source_complete = true;
    return response(transcript);
  });
  confirm.click(); await settle();
  const confirmed = ui.requests.find(request => request.url.endsWith('/transcripts/transcript-1/confirm-complete'));
  assert.deepEqual(confirmed.body, {confirmed: true, notes});
  assert.equal(ui.document.getElementById('cs-source-transcript-1').querySelector('[data-confirm-record]'), null);
  const writerData = clientData(); writerData.sources = [record('transcript-1', 'transcript', {source: 'read_ai', source_complete: false}, {status: 'pending'})];
  const writer = await mount({data: writerData, canReview: false});
  await writer.chooseClient('client-a');
  assert.equal(writer.document.querySelector('[data-confirm-source]'), null, 'Completeness confirmation requires reviewer access');
}

async function testSentCopyRequiresInspectedSanitizedReference() {
  const data = clientData();
  const sent = record('sent-1', 'sent_copy', {title: 'Exact sent copy', channel: 'email', subject: 'Hello Morgan', body: 'Hello Morgan at morgan@example.invalid.'}, {status: 'imported'});
  data.sources = [sent];
  const ui = await mount({data});
  await ui.chooseClient('client-a');
  const card = ui.document.getElementById('cs-source-sent-1');
  assert.equal(card.querySelectorAll('[data-decision]').length, 0, 'Exact sent copy must not be directly approved as a reusable example');
  ui.route('GET', '/api/copy-studio/clients/client-a/records/sent-1', () => response(sent));
  card.querySelector('[data-prepare-reference="sent-1"]').click(); await settle();
  const save = card.querySelectorAll('button').find(button => button.textContent === 'Save reference for review');
  assert.ok(save);
  assert.equal(save.disabled, true, 'Saving a reusable reference requires a sanitization attestation');
  await ui.change('reference-subject-sent-1-0', 'A relevant question');
  await ui.change('reference-body-sent-1-0', 'Hello {{first_name}}, could this offer be relevant?');
  const sanitized = card.querySelector('input[type="checkbox"]');
  sanitized.checked = true; sanitized.dispatchEvent(new UiEvent('change'));
  ui.route('POST', '/api/copy-studio/clients/client-a/records/sent-1/reference', request => {
    const reference = record('reference-1', 'example', {...request.body, channel: 'email'}, {status: 'pending', parent_id: 'sent-1'});
    data.examples.push(reference);
    return response(reference, 201);
  });
  save.click(); await settle();
  const saved = ui.requests.find(request => request.url.endsWith('/records/sent-1/reference'));
  assert.deepEqual(saved.body, {subject: 'A relevant question', body: 'Hello {{first_name}}, could this offer be relevant?', sanitized: true});
  assert.ok(ui.requests.findIndex(request => request.url.endsWith('/records/sent-1') && request.method === 'GET') < ui.requests.indexOf(saved));
  assert.equal(sent.content.body, 'Hello Morgan at morgan@example.invalid.', 'Preparing a reusable reference must preserve exact sent evidence');
}

async function testSavedDraftChecksAndKeylessAssembly() {
  const data = clientData(), ui = await mount({data});
  await ui.chooseClient('client-a');
  const check = ui.get('draft-review').querySelector('[data-check-draft="draft-1"]');
  assert.ok(check && !check.disabled);
  ui.route('POST', '/api/copy-studio/clients/client-a/drafts/draft-1/check', () => response({id: 'semantic-check', status: 'running', client_id: 'client-a'}));
  check.click(); await settle();
  const checked = ui.requests.find(request => request.url.endsWith('/drafts/draft-1/check'));
  assert.deepEqual(checked.body, {});
  const dirty = await mount(); await dirty.chooseClient('client-a');
  const edit = dirty.get('angles').querySelector('[data-copy-path="angles.0.steps.0.body"]');
  edit.value = 'Unsaved changes'; edit.dispatchEvent(new UiEvent('input'));
  assert.equal(dirty.get('draft-review').querySelector('[data-check-draft]').disabled, true,
    'Semantic checks must target a saved immutable version, not unsaved edits');

  const keyless = await mount({modelConfigured: false});
  await keyless.chooseClient('client-a');
  assert.equal(keyless.get('generate').disabled, true, 'Fresh model drafting must be unavailable when the model is unconfigured');
  await keyless.change('mode', 'assemble', 'change');
  const selected = keyless.get('examples').querySelector('input[type="checkbox"]');
  selected.checked = true; selected.dispatchEvent(new UiEvent('change'));
  assert.equal(keyless.get('generate').disabled, false, 'Approved-example assembly must remain usable without a model key');
  keyless.route('POST', '/api/copy-studio/clients/client-a/generate', () => response({id: 'keyless-assemble', status: 'running', client_id: 'client-a'}));
  await keyless.submit('generate-form');
  assert.equal(keyless.requests.find(request => request.url.endsWith('/generate')).body.mode, 'assemble');
  const noLinkedInData = clientData(); noLinkedInData.examples = noLinkedInData.examples.filter(example => example.content.channel !== 'linkedin');
  const missingReference = await mount({data: noLinkedInData}); await missingReference.chooseClient('client-a');
  await missingReference.change('channel', 'linkedin', 'change');
  assert.equal(missingReference.get('generate').disabled, true, 'LinkedIn drafting requires an approved LinkedIn Copy Bank reference');
}

async function testEstimatedCostDisplayAndRefresh() {
  const empty=await mount();
  assert.equal(empty.get('cost-today').textContent,'$0.00');
  assert.equal(empty.get('cost-month').textContent,'$0.00');
  const costs=zeroCosts();
  costs.today={amount:1.25,known_cost:1.25,total_calls:2,unpriced_calls:0,unknown_usage_calls:0};
  costs.month={amount:null,known_cost:3.5,total_calls:4,unpriced_calls:2,unknown_usage_calls:0};
  costs.all_time={amount:null,known_cost:5,total_calls:8,unpriced_calls:1,unknown_usage_calls:1};
  const ui=await mount({costs});
  assert.equal(ui.get('cost-today').textContent,'$1.25');
  assert.match(ui.get('cost-month').textContent,/Unpriced/);
  assert.match(ui.get('cost-month').textContent,/\$3\.50 priced so far/);
  assert.match(ui.get('cost-all-time').textContent,/Incomplete/);
  assert.match(ui.get('cost-all-time').textContent,/\$5\.00 priced so far/);
  ui.route('GET','/api/copy-studio/costs',()=>response(zeroCosts()));
  await ui.click('cost-refresh');
  assert.equal(ui.get('cost-all-time').textContent,'$0.00');
  ui.route('GET','/api/copy-studio/costs',()=>response({error:'Cost data unavailable'},503));
  await ui.click('cost-refresh');
  assert.equal(ui.get('cost-today').textContent,'Unavailable','A failed cost fetch must not imply zero spend');
}

async function testConflictingFactsNeedAnExplicitVersionedDecision() {
  const data=clientData(),content=structuredClone(data.brief.content);
  content.facts.push({...structuredClone(content.facts[0]),id:'old-offer',text:'An older sourced offer.'});
  content.contradictions=[{key:'Core offer',fact_ids:['fact-1','old-offer'],resolution:'Prefer the latest statement.'}];
  const pending=record('brief-conflict','brief',content,{status:'pending',created_at:'2026-10-28T15:00:00Z'});data.pending_briefs=[pending];
  const original=JSON.stringify(pending);
  const ui=await mount({data});await ui.chooseClient('client-a');
  assert.equal(ui.get('brief-review').querySelector('[data-decision="approved"]').disabled,true,
    'A free-text resolution must not allow two contradictory facts to stay active');
  const choice=ui.get('facts').querySelector('[data-fact-decision="1"]');
  assert.equal(choice.checked,true);choice.checked=false;choice.dispatchEvent(new UiEvent('change'));
  assert.equal(ui.get('facts').querySelector('[data-fact-index="1"]').value,'An older sourced offer.',
    'Excluding a fact must preserve its statement and citations for review');
  assert.equal(ui.get('brief-review').querySelector('[data-decision="approved"]').disabled,true,
    'An unsaved exclusion cannot approve the previous immutable version');
  ui.route('POST','/api/copy-studio/clients/client-a/records',request=>{
    const revised=record('brief-decided','brief',{...request.body.content,contradictions:[]},{status:'pending',parent_id:pending.id,created_at:'2026-10-28T16:00:00Z'});
    data.pending_briefs.unshift(revised);return response(revised);
  });
  await ui.submit('brief-form');
  const saved=ui.requests.find(request=>request.method==='POST'&&request.url.endsWith('/records'));
  assert.equal(saved.body.content.facts[1].decision,'exclude');
  assert.equal(saved.body.content.facts[1].text,'An older sourced offer.');
  assert.deepEqual(saved.body.content.facts[1].evidence,content.facts[1].evidence);
  assert.equal(JSON.stringify(pending),original,'Saving fact decisions must not rewrite source versions');
  assert.equal(ui.get('facts').querySelector('[data-fact-decision="1"]').checked,false);
  assert.equal(ui.get('brief-review').querySelector('[data-decision="approved"]').disabled,false);
}

async function testResumeDurableJobAndRefreshCosts() {
  const data=clientData();data.jobs=[{id:'saved-job',kind:'generate',status:'running'}];
  const ui=await mount({data});await ui.chooseClient('client-a');
  assert.equal(ui.get('generate').disabled,true,'Reopening a running task must resume progress instead of allowing duplicate work');
  ui.route('GET','/api/copy-studio/clients/client-a/jobs/saved-job',()=>{
    data.jobs[0].status='succeeded';return response({...data.jobs[0],result:{record_id:'draft-1'}});
  });
  await ui.tick();
  assert.equal(ui.get('generate').disabled,false);
  assert.ok(ui.requests.some(request=>request.url==='/api/copy-studio/costs'),'Completion must refresh estimated costs');
  const partialData=clientData();partialData.jobs=[{id:'partial-sync',kind:'sync',status:'running'}];
  const partial=await mount({data:partialData});await partial.chooseClient('client-a');
  partial.route('GET','/api/copy-studio/clients/client-a/jobs/partial-sync',()=>{
    partialData.jobs[0].status='succeeded';return response({...partialData.jobs[0],result:{complete:false,warnings:['Some source pages could not be read.']}});
  });
  await partial.tick();
  assert.equal(partial.get('status').dataset.tone,'warning');
  assert.match(partial.get('status').textContent,/partial source coverage/);
}

async function main() {
  await testSafeRenderingAndClientIsolation();
  await testGenerationAndPollingStayClientScoped();
  await testImmutableRevisionsAndApprovalGate();
  await testCompletedJobLoadsItsSavedVersion();
  await testPendingBriefEditingAndReviewAreVersionSpecific();
  await testTranscriptInspectionAndCompletenessConfirmation();
  await testSentCopyRequiresInspectedSanitizedReference();
  await testSavedDraftChecksAndKeylessAssembly();
  await testEstimatedCostDisplayAndRefresh();
  await testConflictingFactsNeedAnExplicitVersionedDecision();
  await testResumeDurableJobAndRefreshCosts();
  console.log('Copy Studio UI checks passed: safe source rendering, client isolation, generation contract, polling scope, immutable revisions and approval gates.');
}

main().catch(error => { console.error(error); process.exitCode = 1; });
