import { readFile, writeFile, mkdir, copyFile, readdir } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { rolldown } from 'rolldown';
import { plannedRelease, validateRelease, releaseReady } from './release.mjs';

const site = dirname(fileURLToPath(import.meta.url));
const root = dirname(site), dist = join(site, 'dist');
const pretext = JSON.parse(await readFile(join(root, 'node_modules/@chenglou/pretext/package.json'), 'utf8'));
const lock = JSON.parse(await readFile(join(root, 'package-lock.json'), 'utf8'));
if (pretext.version !== '0.0.8' || lock.packages['node_modules/@chenglou/pretext'].version !== pretext.version)
  throw new Error('The site requires the locked Pretext 0.0.8 dependency.');
let release;
try { release = JSON.parse(await readFile(join(site, 'release.json'), 'utf8')); }
catch (error) { if (error.code !== 'ENOENT') throw error; release = plannedRelease; }
validateRelease(release);
const ready = releaseReady(release);
const escape = value => String(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
const label = text => `<span class="label" data-measure>${escape(text)}</span>`;

function page(kind, body) {
  const download = kind === 'download';
  return `<!doctype html>
<html lang="en" data-measure-state="unavailable">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${download ? 'Download' : 'Description'} · wordplai</title>
  <meta name="description" content="A Microsoft Word add-in for AI editing of locally marked passages.">
  <link rel="icon" href="/wordplai/assets/wordplai.png" type="image/png">
  <link rel="stylesheet" href="/wordplai/assets/style.css">
  <script type="module" src="/wordplai/assets/site.js"></script>
</head>
<body>
  <a class="skip" href="#main">Skip to content</a>
  <div class="shell">
    <header>
      <div class="brand"><img src="/wordplai/assets/wordplai.png" width="64" height="64" alt=""><p>wordplai</p></div>
      <nav class="tabs" aria-label="Pages">
        <a href="/wordplai/"${download ? '' : ' aria-current="page"'}>${label('Description')}</a>
        <a href="/wordplai/download/"${download ? ' aria-current="page"' : ''}>${label('Download')}</a>
      </nav>
    </header>
    <main id="main">${body}</main>
    <footer><span>Early build</span><a href="https://github.com/GeorgeFejer91/wordplai">${label('Source on GitHub')}</a></footer>
  </div>
</body>
</html>
`;
}

const description = `
      <h1>A writing assistant for Word</h1>
      <p>wordplai connects an AI writing assistant to passages you mark in Microsoft Word. You choose which text it may edit and what each color means.</p>
      <h2>Work in your document</h2>
      <ol>
        <li>Mark a passage and assign its color a prompt in Word.</li>
        <li>Connect your writing assistant through MCP and ask it to read the document.</li>
        <li>Ask for changes. Eligible prompt passages are edited directly, with a concise rationale in a Word comment.</li>
      </ol>
      <p>Completed wording edits turn green and become protected. Only you can assign or change a color’s function in Word. AI-added color alone grants no editing permission.</p>
      <p>Existing tracked edits keep their review process. A Discard rule still requires Word review before removing a complete passage.</p>
      <h2>Requirements</h2>
      <p>Windows 10 or 11 x64 (Intel/AMD), desktop Microsoft Word, and .NET Framework 4.8. An AI client that supports MCP is needed to connect a writing assistant.</p>
      <h2>Current limits</h2>
      <p>This is an early, unsigned build. Word’s foreground editing guard can refuse an edit. Real WebGPT edit timing, the document picker’s appearance inside ChatGPT, and physical click and keyboard interactions still need acceptance testing.</p>
      <p>See <a href="/wordplai/download/">Download</a> for the installer and its release details.</p>`;

const action = ready
  ? `<a class="action" id="download-installer" href="${escape(release.downloadUrl)}">${label('Download for Windows')}</a>`
  : `<span class="action" aria-disabled="true">${label('Download unavailable')}</span>`;
const download = `
      <h1>Download wordplai</h1>
      <p>Installer for Windows 10 and 11 x64 (Intel/AMD) with desktop Microsoft Word and .NET Framework 4.8. This preview does not support 32-bit or ARM Windows.</p>
      <p class="download-action">${action}</p>
      ${ready ? '<p class="notice">This preview is unsigned. Windows may show a publisher warning. Review the source and checksum before installing.</p>' : '<p class="notice">The installer’s verified release details are not available yet. The download will appear here once they are published.</p>'}
      <dl class="metadata">
        <dt>Version</dt><dd id="release-version">${escape(release.version)}</dd>
        <dt>File</dt><dd>${escape(release.asset)}</dd>
        <dt>Size</dt><dd id="release-size">${release.sizeBytes === null ? 'Unavailable' : `${release.sizeBytes.toLocaleString('en-US')} bytes`}</dd>
        <dt>Published</dt><dd id="release-published">${release.publishedAt === null ? 'Unavailable' : `<time datetime="${escape(release.publishedAt)}">${escape(release.publishedAt.slice(0, 10))}</time>`}</dd>
        <dt>SHA-256</dt><dd><code class="checksum" id="release-sha256">${release.sha256 === null ? 'Unavailable' : escape(release.sha256)}</code></dd>
      </dl>
      <h2>Before you install</h2>
      <p>Save and close your Word documents before running the installer. The add-in appears in Word after installation. Source code and setup notes are in the <a href="https://github.com/GeorgeFejer91/wordplai#readme">repository README</a>.</p>
      <p class="muted">Early build: foreground edit refusals are possible. Real WebGPT timing, ChatGPT’s picker rendering, and physical interactions remain unverified.</p>`;

await mkdir(join(dist, 'assets'), { recursive: true });
await mkdir(join(dist, 'download'), { recursive: true });
const bundle = await rolldown({ input: join(site, 'site.js') });
try { await bundle.write({ file: join(dist, 'assets/site.js'), format: 'es', minify: true }); }
finally { await bundle.close(); }
await copyFile(join(site, 'style.css'), join(dist, 'assets/style.css'));
for (const file of await readdir(join(site, 'assets'))) await copyFile(join(site, 'assets', file), join(dist, 'assets', file));
await writeFile(join(dist, 'index.html'), page('description', description));
await writeFile(join(dist, 'download/index.html'), page('download', download));
await writeFile(join(dist, 'release.json'), JSON.stringify(release, null, 2) + '\n');
await writeFile(join(dist, '.nojekyll'), '');
const inputs = ['site/build.mjs', 'site/site.js', 'site/style.css', 'site/release.mjs', 'package-lock.json',
  ...((await readdir(join(site, 'assets'))).map(file => `site/assets/${file}`))];
if (release !== plannedRelease) inputs.push('site/release.json');
const hashes = {};
for (const file of inputs) hashes[file] = createHash('sha256').update(await readFile(join(root, file))).digest('hex');
await writeFile(join(dist, 'build-manifest.json'), JSON.stringify({ pretext: pretext.version, releaseReady: ready, inputs: hashes }, null, 2) + '\n');
console.log(`Built Description and Download; installer metadata ${ready ? 'ready' : 'unavailable'}.`);
