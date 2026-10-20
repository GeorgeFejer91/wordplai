import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir, access } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, extname, join, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';
import { plannedRelease, validateRelease, releaseReady } from './release.mjs';
import './build.mjs';

const site = dirname(fileURLToPath(import.meta.url));
const dist = join(site, 'dist'), evidence = join(site, 'test-results');
await mkdir(evidence, { recursive: true });
const release = validateRelease(JSON.parse(await readFile(join(dist, 'release.json'), 'utf8')));
let checks = 0;
function check(condition, message) { assert.ok(condition, message); checks++; }
for (const patch of [{ downloadUrl: 'https://example.com/setup.exe' }, { downloadUrl: 'https://github.com/GeorgeFejer91/wordplai/releases/latest' },
  { sha256: 'unknown' }, { sizeBytes: 0 }, { sizeBytes: 1.5 }, { publishedAt: 'unknown' }, { asset: 'other.exe' }, { extra: true }]) {
  assert.throws(() => validateRelease({ ...plannedRelease, ...patch })); checks++;
}
check(!releaseReady(plannedRelease), 'Missing metadata cannot enable a download.');
const mime = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.json': 'application/json', '.ttf': 'font/ttf', '.png': 'image/png' };
const server = createServer(async (req, res) => {
  try {
    let path = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
    if (!path.startsWith('/wordplai/')) { res.writeHead(404).end(); return; }
    path = path.slice('/wordplai/'.length);
    if (!extname(path)) path += `${path.endsWith('/') || !path ? '' : '/'}index.html`;
    const file = resolve(dist, path);
    if (!file.startsWith(dist + sep)) { res.writeHead(403).end(); return; }
    res.writeHead(200, { 'Content-Type': `${mime[extname(file)] || 'application/octet-stream'}; charset=utf-8` });
    res.end(await readFile(file));
  } catch { res.writeHead(404).end(); }
});
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
const origin = `http://127.0.0.1:${server.address().port}`;
let executablePath;
for (const file of [process.env.CHROME_PATH, 'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe', '/usr/bin/google-chrome', '/usr/bin/chromium'].filter(Boolean)) {
  try { await access(file); executablePath = file; break; } catch { /* Try the next installed browser. */ }
}
if (!executablePath) throw new Error('Set CHROME_PATH to an installed Chrome or Chromium executable.');
let browser;
const rendered = [], errors = [];
async function geometry(page, name, expectFallback = false) {
  const state = await page.evaluate(() => {
    const labels = [...document.querySelectorAll('[data-measure]')].map(element => {
      const range = document.createRange(); range.selectNodeContents(element);
      const box = element.getBoundingClientRect(), style = getComputedStyle(element);
      return { text: element.textContent, status: element.dataset.measureStatus, reason: element.dataset.measureReason,
        metrics: element.dataset.measureMetrics ? JSON.parse(element.dataset.measureMetrics) : null,
        font: style.fontFamily, size: parseFloat(style.fontSize), box: box.toJSON(),
        textBoxes: [...range.getClientRects()].map(rect => rect.toJSON()),
        clipped: element.scrollWidth > element.clientWidth + 1 || element.scrollHeight > element.clientHeight + 1 };
    });
    const fonts = []; document.fonts.forEach(font => fonts.push({ family: font.family, weight: font.weight, status: font.status }));
    const overflow = [...document.querySelectorAll('body *')].filter(element => element.getBoundingClientRect().right > innerWidth + 1)
      .map(element => ({ tag: element.tagName, class: element.className, text: element.textContent.slice(0, 80), right: element.getBoundingClientRect().right }));
    return { width: innerWidth, scrollWidth: document.documentElement.scrollWidth, measureState: document.documentElement.dataset.measureState,
      stacked: document.querySelector('nav').dataset.stacked, labels, fonts, overflow };
  });
  rendered.push({ name, ...state });
  check(state.scrollWidth <= state.width + 1, `${name}: page must not overflow horizontally.`);
  for (const label of state.labels) {
    check(!label.clipped, `${name}: ${label.text} is clipped.`);
    check(label.textBoxes.every(rect => rect.left >= label.box.left - 1 && rect.right <= label.box.right + 1), `${name}: full label must fit its allocated box.`);
    if (expectFallback) check(label.status === 'browser-fallback', `${name}: unsupported measurement must be reported as fallback.`);
    else {
      check(['single-line', 'wrapped'].includes(label.status), `${name}: real Pretext measurement must fit: ${label.status}.`);
      check(label.metrics && label.metrics.lineCount > 0 && label.metrics.maxLineWidth <= label.box.width + 1, `${name}: bounded text metrics are required.`);
      check(label.metrics.lineCount * label.metrics.lineHeight <= label.box.height + 1, `${name}: measured height must fit.`);
    }
  }
  if (!expectFallback) check(state.fonts.filter(font => font.family === 'Wordplai Sans').every(font => font.status === 'loaded'), `${name}: named bundled fonts must actually load.`);
  return state;
}

try {
  browser = await chromium.launch({ executablePath, headless: true });
  const context = await browser.newContext();
  const page = await context.newPage();
  page.on('pageerror', error => errors.push(error.message));
  const cases = [
    { name: 'desktop', width: 1280, height: 900 },
    { name: 'small', width: 320, height: 720 },
    { name: 'landscape', width: 568, height: 320 },
    { name: 'text-200-percent', width: 320, height: 720, enlarge: true },
    { name: 'long-labels', width: 320, height: 720, long: true },
    { name: 'text-spacing', width: 320, height: 720, spacing: true },
  ];
  for (const route of ['', 'download/']) {
    for (const scenario of cases) {
      await page.setViewportSize({ width: scenario.width, height: scenario.height });
      await page.goto(`${origin}/wordplai/${route}`);
      await page.waitForFunction(() => document.documentElement.dataset.measureState === 'ready');
      check(await page.locator('nav a').count() === 2, 'Navigation has exactly two ordinary page links.');
      check(JSON.stringify(await page.locator('nav a').allTextContents()) === JSON.stringify(['Description', 'Download']), 'Navigation labels are exact.');
      check(JSON.stringify(await page.locator('nav a').evaluateAll(links => links.map(link => link.getAttribute('href')))) === JSON.stringify(['/wordplai/', '/wordplai/download/']), 'Both permanent routes are ordinary links.');
      check(await page.locator('nav a[aria-current="page"]').textContent() === (route ? 'Download' : 'Description'), 'Current page is accessible.');
      check(await page.locator('[role="tab"], [role="tablist"]').count() === 0, 'Page navigation is not a fake tab widget.');
      const originalSizes = await page.locator('[data-measure]').evaluateAll(labels => labels.map(label => parseFloat(getComputedStyle(label).fontSize)));
      if (scenario.enlarge) {
        await page.evaluate(() => document.documentElement.style.fontSize = `${parseFloat(getComputedStyle(document.documentElement).fontSize) * 2}px`);
      }
      if (scenario.long) await page.evaluate(() => {
        document.documentElement.lang = 'de';
        const labels = document.querySelectorAll('nav .label');
        labels[0].textContent = 'Beschreibung und Hinweise zur Bearbeitung Ihres Dokuments';
        labels[1].textContent = 'Windows Installationsdatei und veröffentlichte Prüfsumme herunterladen';
        document.querySelector('.action .label')?.replaceChildren('Installationsdatei für Microsoft Word unter Windows herunterladen');
      });
      if (scenario.spacing) await page.addStyleTag({ content: '* { line-height: 1.5 !important; letter-spacing: .12em !important; word-spacing: .16em !important; } p { margin-bottom: 2em !important; }' });
      await page.waitForTimeout(150);
      const state = await geometry(page, `${route || 'description/'}${scenario.name}`, scenario.spacing);
      if (scenario.enlarge) check(state.labels.every((label, index) => label.size >= originalSizes[index] * 2 - .1), 'User-enlarged text is never shrunk to fit.');
      if (scenario.long) check(state.stacked === 'true', 'Long translated labels reflow to one column.');
      if (['desktop', 'small', 'text-200-percent'].includes(scenario.name)) await page.screenshot({ path: join(evidence, `${route ? 'download' : 'description'}-${scenario.name}.png`), fullPage: true });
    }
  }
  await page.goto(`${origin}/wordplai/`);
  await page.keyboard.press('Tab');
  check(await page.locator('.skip').evaluate(link => link === document.activeElement), 'Skip link is first in keyboard order.');
  await page.keyboard.press('Tab');
  check(await page.locator('nav a').first().evaluate(link => link === document.activeElement && getComputedStyle(link).outlineStyle !== 'none'), 'Navigation has visible keyboard focus.');
  await page.locator('nav').getByRole('link', { name: 'Download', exact: true }).click();
  check(new URL(page.url()).pathname === '/wordplai/download/', 'Download link navigates to its permanent subpage.');
  check((await page.locator('#release-version').textContent()) === release.version, 'Rendered version matches metadata.');
  check((await page.locator('#release-sha256').textContent()) === (release.sha256 ?? 'Unavailable'), 'Checksum is exact or explicitly unavailable.');
  check((await page.locator('#release-size').textContent()) === (release.sizeBytes === null ? 'Unavailable' : `${release.sizeBytes.toLocaleString('en-US')} bytes`), 'Exact byte count matches release metadata.');
  if (releaseReady(release)) check(await page.locator('#download-installer').getAttribute('href') === release.downloadUrl, 'Installer uses its explicit preview release URL.');
  else check(await page.locator('#download-installer').count() === 0 && await page.locator('[aria-disabled="true"]').count() === 1, 'Incomplete release metadata cannot create a working download link.');
  check(!(await page.content()).includes('/releases/latest'), 'Preview download never uses releases/latest.');
  check(await page.locator('footer a').getAttribute('href') === 'https://github.com/GeorgeFejer91/wordplai', 'Source link is outside two-page navigation.');
  await context.close();

  const noScript = await browser.newContext({ javaScriptEnabled: false, viewport: { width: 320, height: 720 } });
  const fallback = await noScript.newPage();
  await fallback.goto(`${origin}/wordplai/download/`);
  check(await fallback.locator('h1').textContent() === 'Download wordplai' && await fallback.locator('nav a').count() === 2, 'Full page and navigation work without JavaScript.');
  check(await fallback.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), 'CSS fallback fits a narrow screen.');
  await noScript.close();
  const missingFont = await browser.newContext({ viewport: { width: 320, height: 720 } });
  await missingFont.route('**/*.ttf', route => route.abort());
  const readable = await missingFont.newPage();
  await readable.goto(`${origin}/wordplai/`);
  await readable.waitForFunction(() => document.documentElement.dataset.measureState === 'unavailable' &&
    [...document.querySelectorAll('[data-measure]')].every(label => label.hasAttribute('data-measure-status')));
  await geometry(readable, 'missing-font-readable-fallback', true);
  await missingFont.close();
  check(errors.length === 0, `No page script errors: ${errors.join('; ')}`);
  await writeFile(join(evidence, 'results.json'), JSON.stringify({ status: 'PASS', checks, releaseReady: releaseReady(release), release, browser: await browser.version(), rendered }, null, 2) + '\n');
  console.log(`PASS ${checks}: two routes, release metadata, text measurement, responsive layouts and readable fallbacks.`);
} catch (error) {
  await writeFile(join(evidence, 'results.json'), JSON.stringify({ status: 'FAIL', checks, error: error.stack, rendered, errors }, null, 2) + '\n');
  throw error;
} finally {
  if (browser) await browser.close();
  await new Promise(resolve => server.close(resolve));
}
