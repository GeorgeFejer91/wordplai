import { prepareWithSegments, measureNaturalWidth, measureLineStats, setLocale } from '@chenglou/pretext';

const labels = [...document.querySelectorAll('[data-measure]')];
const cache = new Map();
let pending = false, locale = '';
function queue() {
  if (pending) return;
  pending = true;
  requestAnimationFrame(() => { pending = false; measure(); });
}
function prepared(text, style) {
  const font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
  const spacing = parseFloat(style.letterSpacing) || 0;
  const key = `${font}|${spacing}|${text}`;
  if (!cache.has(key)) {
    if (cache.size >= 16) cache.delete(cache.keys().next().value);
    cache.set(key, prepareWithSegments(text, font, { letterSpacing: spacing, whiteSpace: 'normal', wordBreak: 'normal' }));
  }
  return cache.get(key);
}
function measure() {
  try {
    if (!Intl.Segmenter || !document.fonts.check('18px "Wordplai Sans"')) throw new Error('font-or-segmenter-unavailable');
    const nextLocale = document.documentElement.lang || 'en';
    if (nextLocale !== locale) { setLocale(nextLocale); locale = nextLocale; cache.clear(); }
    const rows = labels.map(element => ({ element, style: getComputedStyle(element), rect: element.getBoundingClientRect() }));
    const nav = document.querySelector('.tabs');
    const budget = (nav.getBoundingClientRect().width - parseFloat(getComputedStyle(nav).columnGap)) / 2;
    const stack = rows.filter(row => row.element.closest('nav')).some(row =>
      measureNaturalWidth(prepared(row.element.textContent, row.style)) + 2 * parseFloat(getComputedStyle(row.element.parentElement).paddingInlineStart) + 2 > budget);
    if (nav.dataset.stacked !== String(stack)) { nav.dataset.stacked = String(stack); queue(); return; }
    for (const { element, style, rect } of rows) {
      if (parseFloat(style.wordSpacing) || style.direction === 'rtl') {
        element.dataset.measureStatus = 'browser-fallback';
        element.dataset.measureReason = 'text-spacing-or-direction';
        continue;
      }
      const text = prepared(element.textContent, style), lineHeight = parseFloat(style.lineHeight);
      const stats = measureLineStats(text, rect.width), naturalWidth = measureNaturalWidth(text);
      const fits = stats.maxLineWidth <= rect.width + 1 && stats.lineCount * lineHeight <= rect.height + 1 &&
        (stats.lineCount !== 1 || naturalWidth <= rect.width + 1);
      element.dataset.measureStatus = fits ? (stats.lineCount === 1 ? 'single-line' : 'wrapped') : 'no-fit';
      element.dataset.measureMetrics = JSON.stringify({ width: rect.width, height: rect.height, naturalWidth, lineCount: stats.lineCount,
        maxLineWidth: stats.maxLineWidth, lineHeight, font: style.fontFamily, fontSize: parseFloat(style.fontSize) });
    }
    document.documentElement.dataset.measureState = 'ready';
  } catch (error) {
    document.documentElement.dataset.measureState = 'unavailable';
    labels.forEach(label => { label.dataset.measureStatus = 'browser-fallback'; label.dataset.measureReason = error.message; });
  }
}
await document.fonts.ready;
await Promise.all([document.fonts.load('400 18px "Wordplai Sans"'), document.fonts.load('700 18px "Wordplai Sans"')]).catch(() => {});
new ResizeObserver(queue).observe(document.querySelector('.shell'));
new MutationObserver(queue).observe(document.documentElement, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ['style', 'lang'] });
document.fonts.addEventListener('loadingdone', queue);
queue();
