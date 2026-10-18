import { prepareWithSegments, measureLineStats, measureNaturalWidth, setLocale } from '@chenglou/pretext';

const picker = document.querySelector('#picker');
const heading = document.querySelector('h1');
const documents = document.querySelector('#documents');
const status = document.querySelector('#status');
const instruction = document.querySelector('#instruction');
const navigation = document.querySelector('#navigation');
const previous = document.querySelector('#previous');
const next = document.querySelector('#next');
const page = document.querySelector('#page');
const retry = document.querySelector('#retry');
const pending = new Map();
const cancellations = new Set();
const prepared = new Map();
let requestId = 0, choiceGeneration = 0, items = [], pageIndex = 0, pageSize = 8, busy = false;
let selected = null, ready = false, measurementQueued = false, maximumHeight = Infinity, lastHeight = 0;

function notify(method, params = {}) {
  if (window.parent !== window) window.parent.postMessage({ jsonrpc: '2.0', method, params }, '*');
}

function request(method, params, generation) {
  const id = `wordplai-picker-${++requestId}`;
  return bounded(() => new Promise((resolve, reject) => {
    pending.set(id, { resolve, reject });
    window.parent.postMessage({ jsonrpc: '2.0', id, method, params }, '*');
  }), generation).finally(() => pending.delete(id));
}

function bounded(callback, generation) {
  let timeout, cancel;
  return new Promise((resolve, reject) => {
    cancel = () => reject(new Error('The request was cancelled.'));
    if (generation !== undefined) cancellations.add(cancel);
    timeout = setTimeout(() => reject(new Error('The connection did not respond within 15 seconds.')), 15000);
    Promise.resolve().then(() => {
      if (generation !== undefined && generation !== choiceGeneration) throw new Error('The request was cancelled.');
      return callback();
    }).then(resolve, reject);
  }).finally(() => { clearTimeout(timeout); cancellations.delete(cancel); });
}

function content(result) {
  if (result?.structuredContent) return result.structuredContent;
  if (result?.selection_required !== undefined || result?.manuscript_id) return result;
  const text = result?.content?.find(part => part.type === 'text')?.text;
  if (text) { try { return JSON.parse(text); } catch { return null; } }
  return null;
}

function nativeMessage(result) {
  const message = content(result)?.error?.message ?? result?.content?.find(part => part.type === 'text')?.text;
  return typeof message === 'string' && message.trim() ? Array.from(message.trim()).slice(0, 500).join('') :
    'Word could not read that document. Check Word and choose again.';
}

function setStatus(text, error = false) {
  status.textContent = text;
  status.dataset.error = String(error);
  queueMeasurement();
}

function setBusy(value) {
  busy = value;
  picker.setAttribute('aria-busy', String(value));
  documents.querySelectorAll('button').forEach(button => { button.disabled = value; });
  previous.disabled = value || pageIndex === 0;
  next.disabled = value || (pageIndex + 1) * pageSize >= items.length;
  retry.disabled = value;
}

function renderPage(focus = false) {
  documents.replaceChildren();
  const counts = new Map();
  for (const item of items) counts.set(item.title, (counts.get(item.title) || 0) + 1);
  for (const item of items.slice(pageIndex * pageSize, (pageIndex + 1) * pageSize)) {
    const row = document.createElement('li');
    const button = document.createElement('button');
    button.type = 'button'; button.className = 'document';
    const title = document.createElement('span');
    title.className = 'title'; title.dataset.measure = ''; title.textContent = item.title;
    button.append(title);
    if (counts.get(item.title) > 1) {
      const identity = document.createElement('span');
      identity.className = 'identity'; identity.dataset.measure = '';
      identity.textContent = `Document ID: ${item.manuscript_id}`;
      button.append(identity);
    }
    button.setAttribute('aria-label', `Use ${item.title}${counts.get(item.title) > 1 ? `, document ID ${item.manuscript_id}` : ''}`);
    button.addEventListener('click', () => choose(item));
    row.append(button); documents.append(row);
  }
  page.textContent = `Page ${pageIndex + 1} of ${Math.max(1, Math.ceil(items.length / pageSize))}`;
  navigation.hidden = items.length <= pageSize;
  setBusy(busy); queueMeasurement();
  if (focus) documents.querySelector('button')?.focus();
}

function receive(result) {
  const data = content(result);
  if (result?.isError || data?.error) {
    setStatus(nativeMessage(result), true);
    return;
  }
  if (!data) return;
  if (Array.isArray(data.manuscripts)) {
    if (data.manuscripts.length > 32) {
      setStatus('Word returned an unavailable document list. Ask wordplai to refresh it.', true); return;
    }
    const seen = new Set();
    for (const item of data.manuscripts) {
      if (typeof item?.manuscript_id !== 'string' || !item.manuscript_id.length || item.manuscript_id.length > 128 ||
          typeof item.title !== 'string' || !item.title.length || seen.has(item.manuscript_id)) {
        items = []; selected = null; documents.replaceChildren(); documents.parentElement.hidden = true;
        navigation.hidden = true; retry.hidden = true;
        heading.textContent = 'Document choice unavailable';
        instruction.textContent = 'No document has been selected.';
        setStatus(seen.has(item?.manuscript_id) ? 'Two open documents have the same review identity. Close the extra copy in Word, then ask wordplai to refresh.' :
          'Word returned an invalid document list. Ask wordplai to refresh it.', true);
        return;
      }
      seen.add(item.manuscript_id);
    }
    items = data.manuscripts;
    if (!items.length || (items.length === 1 && !data.selected_manuscript_id)) {
      selected = null; documents.replaceChildren(); documents.parentElement.hidden = true; navigation.hidden = true; retry.hidden = true;
      heading.textContent = items.length ? 'Word document' : 'No Word documents open';
      instruction.textContent = items.length ? items[0].title : 'Open a document in Word to continue.';
      setStatus(items.length ? 'Continue the writing request in Word.' : 'No document has been selected.');
      return;
    }
    if (data.selection_required === true) {
      selected = null; pageIndex = 0; pageSize = 8;
      heading.textContent = 'Choose a Word document';
      instruction.textContent = 'Choose one document. Writing will stay in it until you ask to switch.';
      retry.hidden = true; documents.parentElement.hidden = false;
      setStatus(`${items.length} open documents.`);
      renderPage(); return;
    }
  }
  if (typeof (data.manuscript_id ?? data.selected_manuscript_id) === 'string' && typeof data.title === 'string') {
    selected = { manuscript_id: data.manuscript_id ?? data.selected_manuscript_id, title: data.title };
    documents.parentElement.hidden = true; navigation.hidden = true;
    heading.textContent = 'Word document';
    instruction.textContent = data.title;
    setStatus(`Ready in ${data.title}.`);
  } else {
    setStatus('The document list changed. Ask wordplai to refresh it.', true);
  }
}

async function continueWriting(generation) {
  if (!selected || busy) return;
  if (generation === undefined) generation = ++choiceGeneration;
  if (generation !== choiceGeneration) return;
  const target = selected;
  setBusy(true); retry.hidden = true;
  setStatus(`Continuing in ${target.title}…`);
  try {
    const text = `Continue the writing request in “${target.title}”.`;
    if (ready) await request('ui/message', { role: 'user', content: [{ type: 'text', text }] }, generation);
    else if (window.openai?.sendFollowUpMessage) await bounded(() => window.openai.sendFollowUpMessage({ prompt: text }), generation);
    else throw new Error('No conversation bridge.');
    if (generation === choiceGeneration) setStatus(`Writing continues in ${target.title}.`);
  } catch {
    if (generation !== choiceGeneration) return;
    setStatus(`Selected ${target.title}, but chat did not resume. Try Continue writing.`, true);
    retry.hidden = false;
  } finally { if (generation === choiceGeneration) { setBusy(false); queueMeasurement(); } }
}

async function choose(item) {
  if (busy) return;
  const generation = ++choiceGeneration;
  setBusy(true); retry.hidden = true;
  setStatus(`Opening ${item.title}…`);
  let successful = false;
  try {
    const arguments_ = { manuscript_id: item.manuscript_id, offset: 0 };
    const result = ready ? await request('tools/call', { name: 'read_manuscript', arguments: arguments_ }, generation) :
      window.openai?.callTool ? await bounded(() => window.openai.callTool('read_manuscript', arguments_), generation) : await request('tools/call', { name: 'read_manuscript', arguments: arguments_ }, generation);
    if (generation !== choiceGeneration) return;
    const data = content(result);
    if (result?.isError || data?.error || data?.selection_required) { receive(result); return; }
    if (data?.manuscript_id !== item.manuscript_id) throw new Error('The document was not selected.');
    receive(result);
    successful = true;
  } catch {
    if (generation !== choiceGeneration) return;
    setStatus('Word could not confirm this document. Check Word and choose again.', true);
  } finally { if (generation === choiceGeneration) setBusy(false); }
  if (successful && generation === choiceGeneration) await continueWriting(generation);
}

function hostContext(context) {
  if (context?.theme === 'dark' || context?.theme === 'light') document.documentElement.style.colorScheme = context.theme;
  if (typeof context?.locale === 'string') document.documentElement.lang = context.locale;
  const limit = context?.containerDimensions?.height ?? context?.containerDimensions?.maxHeight;
  if (typeof limit === 'number' && limit > 0) maximumHeight = limit;
  prepared.clear(); queueMeasurement();
}

function queueMeasurement() {
  if (measurementQueued) return;
  measurementQueued = true;
  requestAnimationFrame(measure);
}

async function measure() {
  measurementQueued = false;
  await document.fonts.ready;
  try {
    setLocale(document.documentElement.lang || 'en');
    const measurements = [...picker.querySelectorAll('[data-measure]')].filter(element => element.getClientRects().length).map(element => {
      const style = getComputedStyle(element), width = element.clientWidth;
      const lineHeight = parseFloat(style.lineHeight), spacing = parseFloat(style.letterSpacing) || 0;
      const font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
      const key = JSON.stringify([element.textContent, font, spacing, document.documentElement.lang]);
      let value = prepared.get(key);
      if (!value) {
        value = prepareWithSegments(element.textContent, font, { letterSpacing: spacing, whiteSpace: 'normal', wordBreak: 'normal' });
        if (prepared.size >= 128) prepared.clear();
        prepared.set(key, value);
      }
      const stats = measureLineStats(value, Math.max(1, width - 2));
      return { element, height: Math.max(1, stats.lineCount) * lineHeight + 2,
        fits: stats.maxLineWidth <= width + 1, lines: stats.lineCount, natural: measureNaturalWidth(value) };
    });
    for (const value of measurements) {
      value.element.style.minHeight = `${value.height}px`;
      value.element.dataset.lines = String(value.lines);
      value.element.dataset.fit = value.fits ? 'measured' : 'browser-wrap';
      value.element.dataset.naturalWidth = String(Math.round(value.natural));
    }
    picker.dataset.measurement = 'pretext';
  } catch { picker.dataset.measurement = 'browser-fallback'; }
  const height = Math.ceil(picker.getBoundingClientRect().height);
  if (!selected && height > maximumHeight && pageSize > 1 && items.length > 1) {
    pageSize--; pageIndex = 0; renderPage(); return;
  }
  picker.dataset.heightFit = height > maximumHeight ? 'needs-more-host-space' : 'fit';
  if (ready && height !== lastHeight) { lastHeight = height; notify('ui/notifications/size-changed', { height, width: Math.ceil(document.documentElement.clientWidth) }); }
}

previous.addEventListener('click', () => { if (pageIndex > 0) { pageIndex--; renderPage(true); } });
next.addEventListener('click', () => { if ((pageIndex + 1) * pageSize < items.length) { pageIndex++; renderPage(true); } });
retry.addEventListener('click', () => continueWriting());
window.addEventListener('message', event => {
  if (event.source !== window.parent || event.data?.jsonrpc !== '2.0') return;
  const message = event.data;
  if (pending.has(message.id)) {
    const entry = pending.get(message.id); pending.delete(message.id);
    if (message.error) entry.reject(new Error('The host declined the request.')); else entry.resolve(message.result);
  } else if (message.method === 'ui/notifications/tool-result') receive(message.params);
  else if (message.method === 'ui/notifications/host-context-changed') hostContext(message.params);
  else if (message.method === 'ui/notifications/tool-cancelled') {
    choiceGeneration++;
    for (const cancel of cancellations) cancel();
    setBusy(false); retry.hidden = !selected;
    setStatus(selected ? 'The request was cancelled. Try Continue writing.' : 'The request was cancelled. Choose a document to try again.', true);
  }
});
window.addEventListener('openai:set_globals', event => { if (event.detail?.globals?.toolOutput) receive(event.detail.globals.toolOutput); });
new ResizeObserver(queueMeasurement).observe(picker);
document.fonts.ready.then(queueMeasurement);
if (window.openai?.toolOutput) { receive(window.openai.toolOutput); setBusy(false); }
request('ui/initialize', { appInfo: { name: 'wordplai document picker', version: '1.0.0' }, appCapabilities: { availableDisplayModes: ['inline'] }, protocolVersion: '2026-01-26' })
  .then(result => { ready = true; hostContext(result?.hostContext); notify('ui/notifications/initialized'); setBusy(false); })
  .catch(() => { setBusy(false); if (!items.length && !selected) setStatus('The Word connection did not respond. Ask wordplai to refresh the document choices.', true); });
