import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';

const source = (await readFile(new URL('./document-picker.js', import.meta.url), 'utf8')).replace(/^import[^\n]+\n/, '').split("previous.addEventListener('click'")[0];
const elements = new Map();
class Element {
  constructor() { this.textContent = ''; this.hidden = false; this.dataset = {}; this.children = []; this.parentElement = { hidden: false }; }
  replaceChildren() { this.children = []; }
  append(child) { this.children.push(child); }
  querySelectorAll() { return []; }
  setAttribute() {}
  addEventListener() {}
}
const sent = [];
const context = vm.createContext({ document: { querySelector(selector) { if (!elements.has(selector)) elements.set(selector, new Element()); return elements.get(selector); }, createElement() { return new Element(); } },
  window: { parent: { postMessage(message) {
    sent.push(message);
    if (message.method === 'ui/message') {
      const entry = vm.runInContext(`pending.get(${JSON.stringify(message.id)})`, context);
      clearTimeout(entry.timeout);
      vm.runInContext(`pending.delete(${JSON.stringify(message.id)})`, context);
      if (!Array.isArray(message.params.content) || message.params.content.some(part => part.type !== 'text' || typeof part.text !== 'string'))
        entry.reject(new Error('The host requires an array of message content blocks.'));
      else entry.resolve({});
    }
  } } }, requestAnimationFrame() {}, setTimeout, clearTimeout });
vm.runInContext(source, context);
const receive = result => vm.runInContext(`receive(${JSON.stringify(result)})`, context);
receive({ selection_required: false, manuscripts: [], title: null });
assert.equal(elements.get('h1').textContent, 'No Word documents open');
assert.equal(elements.get('#documents').parentElement.hidden, true);
assert(!elements.get('#status').textContent.includes('Loading'));
receive({ selection_required: false, manuscripts: [{ manuscript_id: 'one', title: 'Only.docx' }], selected_manuscript_id: null, title: 'Only.docx' });
assert.equal(elements.get('h1').textContent, 'Word document');
assert.equal(elements.get('#instruction').textContent, 'Only.docx');
assert.equal(elements.get('#retry').hidden, true);
assert.deepEqual(sent, [], 'Zero/one list results never start tools or send a follow-up.');
receive({ selection_required: true, manuscripts: [{ manuscript_id: 'copy', title: 'A.docx' }, { manuscript_id: 'copy', title: 'B.docx' }] });
assert.match(elements.get('#status').textContent, /same review identity.*Close the extra copy/);
assert.equal(elements.get('#documents').children.length, 0);
assert.equal(elements.get('#documents').parentElement.hidden, true);
const message = 'The chosen document closed. Reopen it in Word. <img src=x> ' + 'x'.repeat(600);
receive({ isError: true, structuredContent: { error: { code: 'document_unavailable', message } } });
assert.equal(elements.get('#status').textContent, message.slice(0, 500));
assert.equal(elements.get('#status').dataset.error, 'true');
receive({ isError: true, content: [{ type: 'text', text: 'Close the duplicate reviewed copy before continuing.' }] });
assert.equal(elements.get('#status').textContent, 'Close the duplicate reviewed copy before continuing.');
receive({ selection_required: false, manuscripts: [{ manuscript_id: 'chosen', title: 'Chosen.docx' }], selected_manuscript_id: 'chosen', title: 'Chosen.docx' });
assert.equal(elements.get('#status').textContent, 'Ready in Chosen.docx.');
assert.deepEqual(sent, []);
await vm.runInContext('ready = true; continueWriting()', context);
assert(Array.isArray(sent[0].params.content), 'The chat continuation uses standard content blocks.');
assert.equal(sent[0].params.role, 'user');
assert.equal(sent[0].params.content[0].type, 'text');
assert.equal(sent[0].params.content[0].text, 'Continue the writing request in “Chosen.docx”.');
assert.equal(elements.get('#status').textContent, 'Writing continues in Chosen.docx.');
sent.length = 0;
context.document.fonts = { ready: Promise.resolve() };
context.document.documentElement = { clientWidth: 640, lang: 'en' };
context.setLocale = () => {};
elements.get('#picker').getBoundingClientRect = () => ({ height: 120 });
await vm.runInContext('ready = false; measure()', context);
assert.deepEqual(sent, [], 'Uninitialized measurement never publishes a size notification.');
elements.get('#picker').getBoundingClientRect = () => ({ height: 180 });
await vm.runInContext('ready = true; measure()', context);
assert.deepEqual(JSON.parse(JSON.stringify(sent)), [{ jsonrpc: '2.0', method: 'ui/notifications/size-changed', params: { height: 180, width: 640 } }],
  'The first initialized measurement publishes the latest size.');
await vm.runInContext('measure()', context);
assert.equal(sent.length, 1, 'Unchanged size stays deduplicated after initialization.');
console.log('PASS: document states, bounded errors, standard continuation content and initialization-gated size publication. DOM doubles only.');
