import { readFile, readdir, stat } from 'node:fs/promises';
import { extname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const repositoryRoot = fileURLToPath(new URL('../../', import.meta.url));
const required = ['AGENTS.md', 'for-ai/README.md', 'for-ai/PROJECT.md', 'for-ai/SKILLS.md',
  'package.json', 'package-lock.json', 'modules/review/selective-editor-contract.json',
  'modules/word/src/ComAddin.cs', 'modules/word/src/McpEndpoint.cs', 'modules/word/src/Policy.cs',
  'modules/word/build.ps1', 'modules/word/test.ps1', 'modules/word/README.md', 'modules/word/THIRD-PARTY.md',
  'modules/word/packaging/installer.nsi', 'modules/word/packaging/fetch-zrok.ps1',
  'modules/word/packaging/zrok.json', 'modules/word/packaging/zrok-LICENSE',
  'modules/word/assets/document-picker.html', 'modules/word/widget/build-manifest.json',
  'assets/wraiter/icon.svg', 'assets/wraiter/icon16.png', 'assets/wraiter/icon.ico'];
const tooling = { '@chenglou/pretext': '0.0.8', 'playwright-core': '1.62.1', rolldown: '1.2.3' };

export async function architectureErrors(root = repositoryRoot) {
  const errors = [];
  for (const path of required) {
    if (!(await stat(join(root, path)).catch(() => null))?.isFile()) errors.push(`Required standalone asset is missing: ${path}`);
  }
  async function inspectControlPlane(directory) {
    for (const entry of await readdir(directory, { withFileTypes: true })) {
      const path = join(directory, entry.name);
      if (entry.isDirectory()) await inspectControlPlane(path);
      else if (/\.(?:cs|dll|exe|html|png|svg|ico|docx|rtf)$/i.test(extname(path)))
        errors.push(`Product output cannot live in for-ai: ${relative(root, path)}`);
    }
  }
  if ((await stat(join(root, 'for-ai')).catch(() => null))?.isDirectory()) await inspectControlPlane(join(root, 'for-ai'));
  if (errors.length) return errors;
  const pkg = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'));
  const lock = JSON.parse(await readFile(join(root, 'package-lock.json'), 'utf8'));
  if (pkg.name !== 'wordplai' || pkg.license !== 'UNLICENSED' || pkg.private !== true)
    errors.push('Standalone tooling must retain the wordplai identity and explicit unlicensed/private package boundary.');
  if (Object.keys({ ...pkg.dependencies, ...pkg.optionalDependencies }).length)
    errors.push('The Word add-in has no Node runtime dependencies.');
  if (JSON.stringify(Object.entries(pkg.devDependencies || {}).sort()) !== JSON.stringify(Object.entries(tooling).sort()))
    errors.push('Build/test dependencies must match the three exact standalone tool pins.');
  for (const [name, version] of Object.entries(tooling)) {
    if (lock.packages?.[`node_modules/${name}`]?.version !== version)
      errors.push(`The lockfile does not contain the pinned ${name} ${version}.`);
  }
  if (lock.name !== pkg.name || lock.version !== pkg.version || lock.packages?.['']?.name !== pkg.name)
    errors.push('The lockfile root must match package.json.');
  if (Object.values(pkg.scripts || {}).some(command => /\b(?:cargo|tauri|undertree|electron|vite)\b/i.test(command)))
    errors.push('Standalone build/test commands cannot depend on the former application stack.');
  return errors;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const errors = await architectureErrors();
  for (const error of errors) process.stderr.write(`architecture: ${error}\n`);
  if (errors.length) process.exitCode = 1;
  else process.stdout.write('Standalone Word architecture boundaries verified.\n');
}
