/**
 * Bundle size gate.
 *
 * Prints the gzipped size of every published artifact and fails when the ESM
 * *main entry* exceeds `BUDGET_GZIP`. Run after `pnpm build`. The budget exists
 * so the size advertised in the README cannot silently rot.
 *
 * The build is code-split (the streaming reader lives in the separate
 * `sheet-lite/stream` entry point), so "main entry size" is the gzipped total
 * of every module reachable from `dist/index.mjs` through STATIC imports.
 * Dynamically imported chunks (the streaming reader, loaded on demand by
 * `Workbook.openStream`) are reported separately and are not gated — but they
 * are still listed so their cost is visible.
 */
import { readFileSync } from 'node:fs';
import { gzipSync } from 'node:zlib';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/** Gzipped budget for the eager main entry, in bytes. */
const BUDGET_GZIP = 20480; // 20 kB

const dist = join(dirname(fileURLToPath(import.meta.url)), '..', 'dist');
const kib = (bytes) => `${(bytes / 1024).toFixed(2)} kB`;
const gzipLen = (path) => gzipSync(readFileSync(path), { level: 9 }).length;

/** Static `import ... from "./x.mjs"` / `import "./x.mjs"` targets (relative only). */
function staticImports(source) {
  const out = new Set();
  const fromRe = /\bfrom\s*["'](\.\/[^"']+)["']/g;
  const sideRe = /\bimport\s*["'](\.\/[^"']+)["']/g;
  for (const re of [fromRe, sideRe]) {
    let m;
    while ((m = re.exec(source))) out.add(m[1].replace(/^\.\//, ''));
  }
  return [...out];
}

/** Relative chunks pulled in dynamically (`import("./x.mjs")` / `require("./x.cjs")`). */
function dynamicImports(source) {
  const out = new Set();
  const re = /\b(?:import|require)\(\s*["'](\.\/[^"']+)["']\s*\)/g;
  let m;
  while ((m = re.exec(source))) out.add(m[1].replace(/^\.\//, ''));
  return [...out];
}

/** All modules reachable from `entry` via the given edge kind (entry included). */
function closure(entry, pick) {
  const seen = new Set();
  const stack = [entry];
  while (stack.length) {
    const name = stack.pop();
    if (seen.has(name)) continue;
    seen.add(name);
    const source = readFileSync(join(dist, name), 'utf8');
    for (const dep of pick(source)) if (!seen.has(dep)) stack.push(dep);
  }
  return [...seen];
}

const main = closure('index.mjs', staticImports).sort();
const mainGzip = main.reduce((sum, name) => sum + gzipLen(join(dist, name)), 0);
const mainRaw = main.reduce((sum, name) => sum + readFileSync(join(dist, name)).length, 0);

console.log('ESM main entry (static graph):');
for (const name of main) {
  const bytes = readFileSync(join(dist, name));
  console.log(`  ${name.padEnd(34)} ${kib(bytes.length).padStart(10)} raw ${kib(gzipLen(join(dist, name))).padStart(10)} gzip`);
}
console.log(`  ${'TOTAL'.padEnd(34)} ${kib(mainRaw).padStart(10)} raw ${kib(mainGzip).padStart(10)} gzip  ${mainGzip > BUDGET_GZIP ? 'OVER' : 'ok'}`);

// Informational: everything the main entry can pull in at runtime (dynamic
// import of the streaming reader) plus the sibling CJS/types artifacts.
const all = closure('index.mjs', (source) => [...staticImports(source), ...dynamicImports(source)]).sort();
const lazyOnly = all.filter((name) => !main.includes(name));
if (lazyOnly.length > 0) {
  console.log('ESM loaded on demand (sheet-lite/stream subpath / Workbook.openStream — not gated):');
  for (const name of lazyOnly) {
    const bytes = readFileSync(join(dist, name));
    console.log(`  ${name.padEnd(34)} ${kib(bytes.length).padStart(10)} raw ${kib(gzipLen(join(dist, name))).padStart(10)} gzip`);
  }
}
for (const name of ['stream.mjs', 'stream.cjs', 'index.cjs', 'index.d.mts'].filter((name) => {
  try {
    readFileSync(join(dist, name));
    return true;
  } catch {
    return false;
  }
})) {
  const bytes = readFileSync(join(dist, name));
  console.log(`  ${name.padEnd(34)} ${kib(bytes.length).padStart(10)} raw ${kib(gzipLen(join(dist, name))).padStart(10)} gzip  (not gated)`);
}

if (mainGzip > BUDGET_GZIP) {
  console.error(`\nESM main entry exceeds the ${kib(BUDGET_GZIP)} gzip budget.`);
  process.exit(1);
}