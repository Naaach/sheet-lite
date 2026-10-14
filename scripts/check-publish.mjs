#!/usr/bin/env node
/**
 * Pre-publish guard for sheet-lite.
 *
 *   node scripts/check-publish.mjs
 *
 * Verifies — without uploading anything (`npm pack --dry-run`) — that:
 *   1. the package identity and the `files` allowlist in package.json are correct;
 *   2. a fresh build succeeds (`npm run build`);
 *   3. the tarball contains EXACTLY the published surface
 *      (dist/*.mjs|.cjs|.d.mts|.d.cts + sourcemaps, README.md, LICENSE, package.json)
 *      and nothing else (no src/, docs/, docs-site/, playground/, tests/, scripts/, reports/);
 *   4. the exact `name@version` is still free on the npm registry;
 *   5. you are authenticated against registry.npmjs.org.
 *
 * Exit code 0 = safe to publish. Exit code 1 = a blocking check failed.
 *
 * If the build legitimately starts shipping additional dist/ artifacts, update
 * EXPECTED_FILES below deliberately — the strict equality check is the point.
 */

import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import process from 'node:process';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const NPM = 'npm'; // spawned through a shell so npm.cmd also works on Windows

/** The exact, complete set of paths allowed in the published tarball. */
const EXPECTED_FILES = new Set([
  'package.json',
  'README.md',
  'LICENSE',
  'dist/index.mjs',
  'dist/index.mjs.map',
  'dist/index.cjs',
  'dist/index.cjs.map',
  'dist/index.d.mts',
  'dist/index.d.cts',
]);

/** Defense in depth: nothing matching these may ever appear in the tarball. */
const FORBIDDEN = [
  'src/',
  'docs/',
  'docs-site/',
  'playground/',
  'test/',
  'tests/',
  'scripts/',
  'reports/',
  '.npmrc',
  '.env',
];

let blocking = 0;
const ok = (msg) => console.log(`  [ok]   ${msg}`);
const fail = (msg) => {
  blocking += 1;
  console.log(`  [FAIL] ${msg}`);
};
const warn = (msg) => console.log(`  [warn] ${msg}`);
const header = (msg) => console.log(`\n${msg}`);

/** Run npm with args; capture status/stdout/stderr instead of throwing. */
function npm(args) {
  // Single command string (not an args array) so `shell` doesn't trip Node's
  // DEP0190 on Windows; every argument is a compile-time constant below.
  const command = `${NPM} ${args.join(' ')}`;
  try {
    const stdout = execFileSync(command, {
      cwd: ROOT,
      shell: process.platform === 'win32',
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    return { status: 0, stdout, stderr: '' };
  } catch (err) {
    return {
      status: typeof err.status === 'number' ? err.status : 1,
      stdout: String(err.stdout ?? ''),
      stderr: String(err.stderr ?? err.message ?? ''),
    };
  }
}

/** Tolerant JSON parse for `npm pack --json` output (skips any lifecycle noise). */
function parsePackJson(text) {
  try {
    return JSON.parse(text);
  } catch {
    /* fall through */
  }
  const start = text.search(/^\[/m);
  if (start !== -1) {
    try {
      return JSON.parse(text.slice(start));
    } catch {
      /* fall through */
    }
  }
  return null;
}

const pkg = JSON.parse(readFileSync(resolve(ROOT, 'package.json'), 'utf8'));

header(`sheet-lite publish check — ${pkg.name}@${pkg.version}`);

// 1) Package identity + files allowlist -------------------------------------------------
header('1. package.json identity');
if (pkg.name === 'sheet-lite') ok(`name: ${pkg.name}`);
else fail(`unexpected package name: ${pkg.name}`);

if (/^\d+\.\d+\.\d+(?:-[\w.-]+)?$/.test(pkg.version)) ok(`version: ${pkg.version} (semver)`);
else fail(`version is not semver: ${pkg.version}`);

const filesAllow = [...(pkg.files ?? [])].sort();
const filesWant = ['LICENSE', 'README.md', 'dist'].sort();
if (JSON.stringify(filesAllow) === JSON.stringify(filesWant)) ok('files allowlist: dist, README.md, LICENSE');
else fail(`files allowlist is ${JSON.stringify(filesAllow)}, expected ${JSON.stringify(filesWant)}`);

const entryPoints = { main: pkg.main, module: pkg.module, types: pkg.types };
for (const [key, value] of Object.entries(entryPoints)) {
  if (typeof value === 'string' && value.startsWith('./dist/')) ok(`${key}: ${value}`);
  else fail(`${key} must point into ./dist/, got: ${value}`);
}
if (pkg.scripts?.prepublishOnly) ok(`prepublishOnly: ${pkg.scripts.prepublishOnly}`);
else fail('missing prepublishOnly script (should run build + test before publish)');

// 2) Fresh build ------------------------------------------------------------------------
header('2. build (npm run build)');
const build = npm(['run', 'build']);
if (build.status === 0) ok('tsdown build succeeded');
else fail(`build failed:\n${build.stderr.trim().split('\n').slice(-10).join('\n')}`);

// 3) Tarball contents (npm pack --dry-run) ----------------------------------------------
header('3. tarball contents (npm pack --dry-run)');
const pack = npm(['pack', '--dry-run', '--json', '--ignore-scripts']);
const parsed = parsePackJson(pack.stdout);
if (!parsed || !Array.isArray(parsed) || !parsed[0]) {
  fail(`could not parse npm pack --json output:\n${(pack.stderr || pack.stdout).trim().slice(-500)}`);
} else {
  const info = parsed[0];
  const paths = info.files.map((f) => f.path.replace(/^\.\//, ''));
  const actual = new Set(paths);

  for (const p of paths) {
    if (FORBIDDEN.some((m) => p.startsWith(m) || p === m)) fail(`forbidden path in tarball: ${p}`);
  }
  const missing = [...EXPECTED_FILES].filter((p) => !actual.has(p));
  const unexpected = paths.filter((p) => !EXPECTED_FILES.has(p));
  if (missing.length === 0) ok('all 9 expected files present');
  else fail(`missing from tarball: ${missing.join(', ')}`);
  if (unexpected.length === 0) ok('no unexpected files in tarball');
  else fail(`unexpected files in tarball: ${unexpected.join(', ')}`);

  if (info.name === pkg.name && info.version === pkg.version) ok(`pack reports ${info.name}@${info.version}`);
  else fail(`pack reports ${info.name}@${info.version}, expected ${pkg.name}@${pkg.version}`);

  const kb = (info.size / 1024).toFixed(1);
  const ukb = (info.unpackedSize / 1024).toFixed(1);
  if (paths.length === EXPECTED_FILES.size) ok(`entryCount: ${info.entryCount}, packed ${kb} kB, unpacked ${ukb} kB`);
}

// 4) Registry availability of name@version ----------------------------------------------
header('4. registry availability');
const view = npm(['view', `${pkg.name}@${pkg.version}`, 'version']);
if (view.status === 0) {
  fail(`registry already has ${pkg.name}@${view.stdout.trim()} — bump the version, do not republish`);
} else if (view.stderr.includes('E404')) {
  ok(`${pkg.name}@${pkg.version} is free (E404)`);
} else {
  warn(`could not check the registry (offline or proxy?):\n    ${view.stderr.trim().split('\n')[0]}`);
}

// 5) Auth -------------------------------------------------------------------------------
header('5. npm authentication');
const who = npm(['whoami']);
if (who.status === 0) ok(`authenticated as ${who.stdout.trim()}`);
else {
  const code = /E(NEEDAUTH|AUTH[^ ]*)/.exec(who.stderr)?.[0] ?? 'not logged in';
  fail(`${code} — publish via the "Publish to npm" GitHub workflow (recommended) or run "npm login" first. See docs/PUBLISHING.md.`);
}

// Summary -------------------------------------------------------------------------------
header(blocking === 0 ? 'RESULT: safe to publish ✔' : `RESULT: ${blocking} blocking check(s) failed ✘`);
process.exitCode = blocking === 0 ? 0 : 1;
