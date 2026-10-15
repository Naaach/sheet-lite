/**
 * Bundle size gate.
 *
 * Prints the gzipped size of every published artifact and fails when the ESM
 * *main entry* exceeds `BUDGET_GZIP`. Run after `pnpm build`. The budget exists
 * so the size advertised in the README cannot silently rot.
 *
 * The build is code-split (the streaming reader and the incremental writer have
 * their own entry points, `sheet-lite/stream` and `sheet-lite/write`), so "main
 * entry size" is the gzipped total of every module reachable from
 * `dist/index.mjs` through STATIC imports. Dynamically imported chunks — the
 * streaming reader behind `Workbook.openStream`, and the writer behind the lazy
 * `Workbook.create()` — are reported separately and are not gated, but they are
 * still listed so their cost is visible instead of invisible.
 *
 * ## Reproducibility
 *
 * tsdown names code-split chunks `<name>-<8-char content hash>.<ext>`, and that
 * hash is written into the *importing* module's source text. Measured raw, the
 * hash is a build artefact, not a size signal: it makes an otherwise identical
 * rebuild report a different number, which is how the same library was measured
 * at 19.59 kB on one build and 19.85 kB on the next — a 0.26 kB swing that can
 * fail a 20 kB budget for reasons that have nothing to do with the code.
 *
 * Every chunk name and every byte read here is therefore normalised
 * (`-<hash>` -> `-HASH`) *before* it is measured or printed, so the gate is
 * reproducible run to run. Content hashes are preserved in the build output;
 * only what this script measures is normalised.
 */
import { readFileSync } from 'node:fs';
import { gzipSync } from 'node:zlib';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Gzipped budget for the eager main entry, in bytes.
 *
 * Raised from 20 kB to 22 kB for the write milestone: the main entry gained
 * `removeWorksheet`'s orphan-part pruning (a synchronous feature, so it cannot
 * live behind the lazy writer chunk), the `.xls`/`.xlsb` container detection and
 * the shared cell/style helpers the two writers now use. The incremental writer
 * itself still ships out of band: adding it only moves ~0.3 kB into the `index`
 * graph (its own chunk is `dist/write.mjs`, loaded on demand by
 * `Workbook.create()`, and is not gated).
 */
const BUDGET_GZIP = 22528; // 22 kB

const dist = join(dirname(fileURLToPath(import.meta.url)), '..', 'dist');
const kib = (bytes) => `${(bytes / 1024).toFixed(2)} kB`;

/**
 * tsdown's content hash in chunk file names (`streaming-B4pQhjuX.mjs`,
 * `value-C7CBD-Ou.mjs`, ...). Rewritten to a fixed `HASH` everywhere below.
 *
 * The character class includes `-` and `_` because base64url hashes contain
 * them (`refs-C_vBFkQB.mjs`), and the leading `-` is anchored by the extension
 * lookahead so a *file* name that merely ends in a hash is not mistaken for one.
 */
const CONTENT_HASH = /-[A-Za-z0-9_-]{8}(?=\.[cm]js\b)/g;

/** Stable display name for a chunk: `value-C7CBD-Ou.mjs` -> `value-HASH.mjs`. */
const stableName = (name) => name.replace(CONTENT_HASH, '-HASH');

/**
 * A chunk's bytes with every embedded content hash rewritten. `latin1` is a
 * lossless byte<->code-unit mapping, so the hash characters are the only bytes
 * that change; gzip then measures code, not a build fingerprint.
 */
const stableBytes = (path) =>
  Buffer.from(readFileSync(path).toString('latin1').replace(CONTENT_HASH, '-HASH'), 'latin1');

/** Raw source, read verbatim: import specifiers must match the files on disk. */
const sourceOf = (name) => readFileSync(join(dist, name), 'utf8');

/** True for a statically known relative specifier (a chunk name we can follow). */
const isChunkName = (spec) => spec.startsWith('./') && !spec.includes('${');

/** Static `import ... from "./x.mjs"` / `import "./x.mjs"` targets (relative only). */
function staticImports(source) {
  const out = new Set();
  const fromRe = /\bfrom\s*[`'"]([^`'"]+)[`'"]/g;
  const sideRe = /\bimport\s*[`'"]([^`'"]+)[`'"]/g;
  for (const re of [fromRe, sideRe]) {
    let m;
    while ((m = re.exec(source))) {
      if (isChunkName(m[1])) out.add(m[1].replace(/^\.\//, ''));
    }
  }
  return [...out];
}

/**
 * Relative chunks pulled in dynamically. tsdown emits these as template
 * literals — ``import(`./streaming-B4pQhjuX.mjs`)`` — so the backtick has to be
 * accepted here; with a `'"/"`-only class the whole on-demand graph was silently
 * dropped and the writer never showed up in this report at all.
 */
function dynamicImports(source) {
  const out = new Set();
  const re = /\b(?:import|require)\(\s*[`'"]([^`'"]+)[`'"]\s*\)/g;
  let m;
  while ((m = re.exec(source))) {
    if (isChunkName(m[1])) out.add(m[1].replace(/^\.\//, ''));
  }
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
    const source = sourceOf(name);
    for (const dep of pick(source)) if (!seen.has(dep)) stack.push(dep);
  }
  return [...seen];
}

/** Normalised name + raw/gzip sizes of one file. */
const measure = (name) => {
  const bytes = stableBytes(join(dist, name));
  return { name: stableName(name), raw: bytes.length, gzip: gzipSync(bytes, { level: 9 }).length };
};

const row = ({ name, raw, gzip }) =>
  `  ${name.padEnd(34)} ${kib(raw).padStart(10)} raw ${kib(gzip).padStart(10)} gzip`;

const sum = (rows, key) => rows.reduce((total, row) => total + row[key], 0);

// --- Gated: the eager main entry, reachable through static imports only. ------
const main = closure('index.mjs', staticImports).sort().map(measure);
const mainGzip = sum(main, 'gzip');
const mainRaw = sum(main, 'raw');

console.log('ESM main entry (static graph) — GATED:');
for (const entry of main) console.log(row(entry));
console.log(
  `  ${'TOTAL'.padEnd(34)} ${kib(mainRaw).padStart(10)} raw ${kib(mainGzip).padStart(10)} gzip  ` +
    `${mainGzip > BUDGET_GZIP ? 'OVER' : 'ok'} (budget ${kib(BUDGET_GZIP)})`,
);

// --- Informational: chunks the main entry loads at runtime instead. -----------
// Everything reachable from `dist/index.mjs` through static *and* dynamic edges,
// minus what is already in the gated graph: the streaming reader and the lazy
// `Workbook.create()` writer.
const all = closure('index.mjs', (source) => [...staticImports(source), ...dynamicImports(source)]).sort();
const gatedNames = new Set(main.map((entry) => entry.name));
const lazyOnly = all.filter((name) => !gatedNames.has(stableName(name))).map(measure);
if (lazyOnly.length > 0) {
  console.log('ESM loaded on demand (Workbook.openStream / lazy Workbook.create) — not gated:');
  for (const entry of lazyOnly) console.log(row(entry));
  console.log(
    `  ${'TOTAL'.padEnd(34)} ${kib(sum(lazyOnly, 'raw')).padStart(10)} raw ${kib(sum(lazyOnly, 'gzip')).padStart(10)} gzip  (not gated)`,
  );
}

// --- Informational: the sibling CJS / types artifacts of the same build. ------
const siblings = ['stream.mjs', 'stream.cjs', 'index.cjs', 'index.d.mts'].filter((name) => {
  try {
    readFileSync(join(dist, name));
    return true;
  } catch {
    return false;
  }
});
if (siblings.length > 0) {
  console.log('Sibling artifacts — not gated:');
  for (const name of siblings) console.log(row(measure(name)));
}

if (mainGzip > BUDGET_GZIP) {
  console.error(`\nESM main entry exceeds the ${kib(BUDGET_GZIP)} gzip budget.`);
  process.exit(1);
}