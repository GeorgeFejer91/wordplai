import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { rolldown } from 'rolldown';

const directory = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(directory, '../../..');
const readJson = async file => JSON.parse(await readFile(file, 'utf8'));
const lock = await readJson(path.join(root, 'package-lock.json'));
const pretext = await readJson(path.join(root, 'node_modules/@chenglou/pretext/package.json'));
if (lock.packages['node_modules/@chenglou/pretext'].version !== pretext.version) throw new Error('Installed Pretext does not match the repository lock.');
const bundle = await rolldown({ input: path.join(directory, 'document-picker.js') });
const { output } = await bundle.generate({ format: 'iife', minify: true });
await bundle.close();
if (output.length !== 1 || output[0].type !== 'chunk') throw new Error('The picker must be one inline bundle.');
const template = await readFile(path.join(directory, 'document-picker.template.html'), 'utf8');
const html = template.replace('/* WORDPLAI_SCRIPT */', output[0].code.replace(/<\/script/gi, '<\\/script'));
const destination = path.join(root, 'modules/word/assets/document-picker.html');
await mkdir(path.dirname(destination), { recursive: true });
await writeFile(destination, html, 'utf8');
const inputs = [...new Set([path.join(directory, 'build.mjs'), path.join(directory, 'document-picker.template.html'),
  path.join(root, 'package-lock.json'), ...Object.keys(output[0].modules)])].sort();
const hash = value => createHash('sha256').update(value).digest('hex');
const manifest = { pretext: pretext.version, output: path.relative(root, destination).replaceAll('\\', '/'),
  sha256: hash(html), bytes: Buffer.byteLength(html), inputs: await Promise.all(inputs.map(async file => ({
    path: path.relative(root, file).replaceAll('\\', '/'), sha256: hash(await readFile(file)) }))) };
await writeFile(path.join(directory, 'build-manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
console.log(JSON.stringify({ output: manifest.output, bytes: manifest.bytes, sha256: manifest.sha256, pretext: pretext.version }));
