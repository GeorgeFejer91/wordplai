import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { chromium } from 'playwright-core';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const directory = path.dirname(fileURLToPath(import.meta.url));
const output = path.resolve(process.argv[2] || 'artifacts/word/document-picker-ui');
await mkdir(output, { recursive: true });
const html = await readFile(path.join(directory, '../assets/document-picker.html'), 'utf8');
assert(!/<(?:script|link|img)[^>]+(?:src|href)\s*=/i.test(html), 'No external resources.');
const fixtureBridge = `<script>
  const legacyReplies = new Map(); let legacyRequestId = 0;
  addEventListener('message', event => {
    if(event.source!==parent||event.data?.jsonrpc!=='2.0')return;
    const message=event.data;
    if(message.method==='fixture/barrier')parent.postMessage({jsonrpc:'2.0',method:'fixture/barrier-ack',params:message.params},'*');
    const reply=legacyReplies.get(message.id);if(!reply)return;legacyReplies.delete(message.id);
    if(message.error)reply.reject(new Error(message.error.message));else reply.resolve(message.result);
  });
  if(location.search.includes('legacy')){
    const request=(method,params)=>new Promise((resolve,reject)=>{const id='legacy-picker-'+(++legacyRequestId);legacyReplies.set(id,{resolve,reject});parent.postMessage({jsonrpc:'2.0',id,method,params},'*');});
    window.openai={callTool:(name,args)=>request('legacy/tools/call',{name,arguments:args}),sendFollowUpMessage:({prompt})=>request('legacy/message',{role:'user',content:[{type:'text',text:prompt}]})};
  }
</script>`;
const server = createServer((request, response) => {
  response.setHeader('Content-Type', 'text/html; charset=utf-8');
  if (request.url.startsWith('/picker')) return response.end(html.replace('<script>', fixtureBridge + '<script>'));
  response.end(`<!doctype html><html><body style="margin:0"><iframe title="Word documents" src="/picker${request.url.includes('legacy') ? '?legacy=1' : ''}" style="width:100%;border:0" sandbox="allow-scripts allow-same-origin"></iframe><script>
  window.fixture={calls:[],messages:[],heldCount:0,barrier:0,legacy:${request.url.includes('legacy')},mode:'success',maximumHeight:10000,items:Array.from({length:32},(_,i)=>({manuscript_id:i.toString(16).padStart(32,'0'),title:i===0||i===1?'Draft.docx':i===2?'<img src=x onerror=alert(1)>.docx':i===3?'Unbroken'+('long'.repeat(34))+'.docx':i===4?'Überarbeiteter Entwurf — 日本語 🧠.docx':'Research document '+(i+1)+'.docx'}))};
  const held=[];let barrier=0;
  window.releaseHeld=(index,error=false)=>held[index](error);
  window.flushPicker=()=>{const serial=++barrier;document.querySelector('iframe').contentWindow.postMessage({jsonrpc:'2.0',method:'fixture/barrier',params:{serial}},'*');return serial;};
  window.deliver=()=>document.querySelector('iframe').contentWindow.postMessage({jsonrpc:'2.0',method:'ui/notifications/tool-result',params:{structuredContent:{selection_required:true,manuscripts:fixture.items}}},'*');
  addEventListener('message',event=>{const m=event.data;if(m?.jsonrpc!=='2.0')return;const answer=result=>event.source.postMessage({jsonrpc:'2.0',id:m.id,result},'*');
  const decline=()=>event.source.postMessage({jsonrpc:'2.0',id:m.id,error:{code:-32000,message:'Synthetic refusal'}},'*');
  const finish=result=>{if(fixture.mode==='hold-read'&&m.method.endsWith('tools/call')||fixture.mode==='hold-message'&&m.method.endsWith('message')){held.push(error=>error?decline():answer(result));fixture.heldCount=held.length;}else answer(result);};
  if(m.method==='fixture/barrier-ack')fixture.barrier=m.params.serial;
  else if(m.method==='ui/initialize'){if(fixture.legacy){decline();deliver();}else answer({hostContext:{theme:'light',containerDimensions:{maxHeight:fixture.maximumHeight}}});}
  else if(m.method==='ui/notifications/initialized')deliver();
  else if(m.method==='ui/notifications/size-changed')document.querySelector('iframe').style.height=m.params.height+'px';
  else if(m.method==='tools/call'||m.method==='legacy/tools/call'){fixture.calls.push(m.params);if(fixture.mode==='timeout')return;const item=fixture.items.find(x=>x.manuscript_id===m.params.arguments.manuscript_id);finish({structuredContent:{...item,text:'Synthetic writing context.',complete:true,context_id:'synthetic-context',revision:'synthetic-revision'}});}
  else if(m.method==='ui/message'||m.method==='legacy/message'){fixture.messages.push(m.params);if(fixture.mode==='message-failure'||!Array.isArray(m.params.content)||m.params.role!=='user')decline();else finish({});}
  });</script></body></html>`);
});
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
let browser;
const observations = [];
const cancellation = [];
async function flushPicker(page) {
  const serial = await page.evaluate(() => flushPicker());
  await page.waitForFunction(serial => fixture.barrier === serial, serial);
}
async function cancelPicker(page, frame, hasSelection = false) {
  await page.evaluate(() => document.querySelector('iframe').contentWindow.postMessage({ jsonrpc: '2.0', method: 'ui/notifications/tool-cancelled', params: {} }, '*'));
  await frame.locator('#status').filter({ hasText: hasSelection ? 'The request was cancelled. Try Continue writing.' : 'The request was cancelled. Choose a document to try again.' }).waitFor();
  await flushPicker(page);
}
async function cancellationCase(legacy, scenario, run) {
  const page = await browser.newPage({ viewport: { width: 800, height: 900 } });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  const label = `${legacy ? 'legacy' : 'standard'}-${scenario}`;
  try {
    await page.clock.install();
    await page.goto(`http://127.0.0.1:${server.address().port}/${legacy ? '?legacy=1' : ''}`);
    const frame = page.frames().find(frame => new URL(frame.url()).pathname === '/picker');
    await frame.locator('#documents button').first().waitFor();
    await page.evaluate(() => {
      fixture.items = [{ manuscript_id: 'cancel-a', title: 'Cancel A.docx' }, { manuscript_id: 'cancel-b', title: 'Cancel B.docx' }];
      deliver();
    });
    await frame.locator('#documents button').filter({ hasText: 'Cancel A.docx' }).waitFor();
    await run(page, frame);
    assert.deepEqual(errors, []);
    cancellation.push({ scenario: label, pass: true });
  } catch (error) {
    await page.screenshot({ path: path.join(output, `${label}-failure.png`), fullPage: true });
    await writeFile(path.join(output, `${label}-failure.html`), await page.content());
    cancellation.push({ scenario: label, pass: false, error: error.message, pageErrors: errors });
  } finally { await page.close(); }
}
async function cancellationScenarios() {
  for (const legacy of [false, true]) {
    for (const staleError of [false, true]) {
      await cancellationCase(legacy, `cancel-a-late-${staleError ? 'error' : 'success'}-then-b`, async (page, frame) => {
        await page.evaluate(() => { fixture.mode = 'hold-read'; });
        await frame.locator('#documents button').first().click();
        await page.waitForFunction(() => fixture.heldCount === 1);
        await cancelPicker(page, frame);
        await frame.locator('#documents button').last().click();
        await page.waitForFunction(() => fixture.heldCount === 2);
        const before = await frame.locator('#status').textContent();
        await page.evaluate(staleError => releaseHeld(0, staleError), staleError);
        await flushPicker(page);
        assert.equal(await frame.locator('#picker').getAttribute('aria-busy'), 'true', 'Cancelled A cannot release B’s busy state.');
        assert.equal(await frame.locator('#status').textContent(), before, 'Cancelled A cannot replace B’s status.');
        assert.equal(await frame.locator('#documents').isVisible(), true, 'Cancelled A cannot select itself.');
        assert.deepEqual(await page.evaluate(() => fixture.messages), [], 'Cancelled A emits no continuation.');
        await page.evaluate(() => { fixture.mode = 'success'; releaseHeld(1); });
        await frame.locator('#status').filter({ hasText: 'Writing continues in Cancel B.docx.' }).waitFor();
        assert.deepEqual(await page.evaluate(() => fixture.messages), [{ role: 'user', content: [{ type: 'text', text: 'Continue the writing request in “Cancel B.docx”.' }] }]);
        assert.equal(await frame.locator('#instruction').textContent(), 'Cancel B.docx');
        await page.clock.fastForward(15001);
        await flushPicker(page);
        assert.equal(await frame.locator('#status').textContent(), 'Writing continues in Cancel B.docx.', 'Cancelled A’s timeout cannot alter B.');
      });
    }
    for (const retry of [false, true]) {
      for (const staleError of [false, true]) {
        await cancellationCase(legacy, `cancel-${retry ? 'retried-' : ''}continuation-late-${staleError ? 'error' : 'success'}`, async (page, frame) => {
          await page.evaluate(retry => { fixture.mode = retry ? 'message-failure' : 'hold-message'; }, retry);
          await frame.locator('#documents button').first().click();
          if (retry) {
            await frame.locator('#retry').waitFor({ state: 'visible' });
            await page.evaluate(() => { fixture.mode = 'hold-message'; });
            await frame.locator('#retry').click();
          }
          await page.waitForFunction(() => fixture.heldCount === 1);
          await cancelPicker(page, frame, true);
          const before = await frame.locator('#status').textContent();
          const messages = await page.evaluate(() => fixture.messages.length);
          await page.evaluate(staleError => releaseHeld(0, staleError), staleError);
          await flushPicker(page);
          assert.equal(await frame.locator('#status').textContent(), before, 'A cancelled continuation acknowledgement cannot claim success.');
          assert.equal(await frame.locator('#picker').getAttribute('aria-busy'), 'false');
          assert.equal(await page.evaluate(() => fixture.messages.length), messages, 'Cancellation never sends another continuation.');
          assert.equal(await frame.locator('#retry').isVisible(), true, 'The selected document allows an explicit continuation retry.');
          await page.evaluate(() => { fixture.mode = 'success'; });
          await frame.locator('#retry').click();
          await frame.locator('#status').filter({ hasText: 'Writing continues in Cancel A.docx.' }).waitFor();
          assert.equal(await page.evaluate(() => fixture.calls.length), 1, 'Explicit continuation retry never rereads the document.');
          assert.equal(await page.evaluate(() => fixture.messages.length), messages + 1);
        });
      }
    }
  }
  await writeFile(path.join(output, 'cancellation-cases.json'), JSON.stringify(cancellation, null, 2));
  assert.deepEqual(cancellation.filter(result => !result.pass), [], 'Every cancellation scenario must preserve the current choice.');
}
try {
  browser = await chromium.launch({ executablePath: process.env.WORDPLAI_CHROME || 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', headless: true });
  await cancellationScenarios();
  const page = await browser.newPage({ viewport: { width: 800, height: 900 } });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.clock.install();
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  const frame = page.frames().find(frame => frame.url().endsWith('/picker'));
  await frame.locator('#documents button').first().waitFor();
  await frame.waitForFunction(() => document.querySelector('#picker').dataset.measurement === 'pretext');
  assert.equal(await frame.locator('#documents button').count(), 8);
  assert.equal(await frame.locator('.identity').count(), 2);
  assert.equal(await frame.locator('img').count(), 0, 'Titles remain DOM text.');
  assert.match(await frame.locator('#documents button').first().getAttribute('aria-label'), /document ID 00000000000000000000000000000000/);
  for (const width of [320, 480, 800, 1200]) {
    await page.setViewportSize({ width, height: 900 });
    await page.clock.runFor(100);
    const fit = await frame.evaluate(() => ({ width: innerWidth, horizontalOverflow: document.documentElement.scrollWidth > innerWidth + 1,
      internalScroll: [...document.querySelectorAll('main,ul,nav,button,[data-measure]')].some(el => el.scrollWidth > el.clientWidth + 2 || el.scrollHeight > el.clientHeight + 2),
      measured: document.querySelectorAll('[data-fit="measured"]').length, height: document.querySelector('main').getBoundingClientRect().height }));
    assert.equal(fit.horizontalOverflow, false, `No horizontal overflow at ${width}.`);
    assert.equal(fit.internalScroll, false, `No internal clipping at ${width}.`);
    assert(fit.measured > 10); observations.push(fit);
  }
  await page.setViewportSize({ width: 320, height: 900 });
  await frame.addStyleTag({ content: 'html { font-size: 32px !important; } [data-measure] { line-height: 1.5 !important; letter-spacing: .12em !important; word-spacing: .16em !important; } p { margin-block-end: 2em !important; }' });
  await page.clock.runFor(100);
  assert.equal(await frame.evaluate(() => document.documentElement.scrollWidth > innerWidth + 1), false, '200% text and user spacing reflow.');
  assert.equal(await frame.evaluate(() => {
    const boxes = [...document.querySelector('#navigation').children].map(element => element.getBoundingClientRect());
    return boxes.some((a, index) => boxes.slice(index + 1).some(b => Math.min(a.right, b.right) - Math.max(a.left, b.left) > 1 && Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top) > 1));
  }), false, 'Enlarged pagination labels never overlap.');
  await page.screenshot({ path: path.join(output, 'picker-320-text200.png'), fullPage: true });
  await frame.locator('style').last().evaluate(element => element.remove());
  await page.setViewportSize({ width: 800, height: 900 });
  await frame.locator('#next').click(); await frame.locator('#next').click(); await frame.locator('#next').click();
  assert.equal(await frame.locator('#page').textContent(), 'Page 4 of 4');
  await frame.locator('#documents button').last().focus();
  await page.keyboard.press('Enter');
  await frame.locator('#status').filter({ hasText: 'Writing continues in Research document 32.docx.' }).waitFor();
  const chosen = await page.evaluate(() => fixture);
  assert.equal(chosen.calls.length, 1); assert.equal(chosen.calls[0].name, 'read_manuscript');
  assert.deepEqual(chosen.calls[0].arguments, { manuscript_id: '0000000000000000000000000000001f', offset: 0 });
  assert.equal(chosen.messages.length, 1);
  assert.deepEqual(chosen.messages[0], { role: 'user', content: [{ type: 'text', text: 'Continue the writing request in “Research document 32.docx”.' }] });
  await page.evaluate(() => { fixture.mode = 'timeout'; deliver(); });
  await frame.locator('#documents button').first().click();
  await page.clock.fastForward(15001);
  await frame.locator('#status[data-error="true"]').waitFor();
  assert.equal(await frame.locator('#documents button').first().isDisabled(), false, 'Timeout permits another choice.');
  await page.evaluate(() => { fixture.mode = 'message-failure'; deliver(); });
  await frame.locator('#documents button').first().click();
  await frame.locator('#retry').waitFor({ state: 'visible' });
  assert.equal(await frame.locator('#retry').isDisabled(), false);
  await page.evaluate(() => { fixture.mode = 'success'; });
  const callsBeforeRetry = await page.evaluate(() => fixture.calls.length);
  await frame.locator('#retry').click();
  await frame.locator('#status').filter({ hasText: 'Writing continues in Draft.docx.' }).waitFor();
  assert.equal(await page.evaluate(() => fixture.calls.length), callsBeforeRetry, 'Continuation retry does not reread the document.');
  await page.evaluate(() => { deliver(); document.querySelector('iframe').contentWindow.postMessage({jsonrpc:'2.0',method:'ui/notifications/host-context-changed',params:{containerDimensions:{maxHeight:420}}},'*'); });
  await page.clock.runFor(500);
  assert.equal(await frame.evaluate(() => document.querySelector('main').dataset.heightFit), 'fit', 'Constrained host adapts pagination.');
  await page.screenshot({ path: path.join(output, 'picker-desktop.png'), fullPage: true });
  assert.deepEqual(errors, []);
  const receipt = { pass: true, documents: 32, maximumPerPage: 8, observations, cancellation, errors, htmlBytes: Buffer.byteLength(html) };
  await writeFile(path.join(output, 'widget-test.json'), JSON.stringify(receipt, null, 2));
  console.log(JSON.stringify(receipt));
} finally { if (browser) await browser.close(); await new Promise(resolve => server.close(resolve)); }
