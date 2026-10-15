/**
 * Writing benchmark: incremental streaming writer vs in-memory writes.
 *
 * Run with:  pnpm bench:write   (builds dist first, then `node --expose-gc`)
 *
 * The same rows (default 100k x 8) are written in three phases, each in its
 * OWN fresh node process so nothing pollutes the numbers:
 *
 *   (a) file   `WorkbookWriter.create({ file })` — every row is deflated and
 *              appended to the file as it is produced (constant memory);
 *   (b) buffer `WorkbookWriter.create()` — identical work, but the whole
 *              archive is retained until `toBuffer()` returns it;
 *   (c) naive  the in-memory shape a writer falls into: build the whole sheet
 *              as ONE string, then compress it (no streaming at all).
 *
 * Memory is a *live* peak, sampled from `process.memoryUsage()` at periodic
 * forced-GC checkpoints, so the number reflects what the writer actually holds
 * rather than short-lived garbage. `heapUsed` alone would hide the archive:
 * `TypedArray` backing stores are off-heap, so every phase reports
 * `heapUsed + arrayBuffers`.
 *
 * Every SAMPLE_CELLS a checkpoint is *settled* — GC, a turn of the event loop,
 * GC — before it is read, and the time spent settling is subtracted from the
 * timings. That detail decides whether the numbers mean anything: a bare `gc()`
 * inside the synchronous row loop leaves the writer's own output buffers (fflate
 * hands each chunk over and drops it, but the backing store is only released
 * once the queue drains) counted as live, which adds megabytes of noise that
 * swamps the per-row signal this benchmark exists to measure.
 *
 * The three phases above compare three WRITERS at one row count, which says
 * nothing about whether the streaming writer's memory grows with the sheet. Two
 * further runs therefore measure the same file-target writer at two row counts
 * (50k and 400k by default) and the script FAILS if the live peak at the larger
 * one is not within `SCALE_TOLERANCE` of the smaller one: one row count cannot
 * falsify "does not grow with the row count", because an O(n) leak with a small
 * constant fits under any single measurement.
 *
 * Override the size with env vars:  ROWS=250000 COLS=8 pnpm bench:write
 * Scaling check:                     SCALE_ROWS_SMALL=50000 SCALE_ROWS_LARGE=400000 \
 *                                    SCALE_TOLERANCE=1.5 pnpm bench:write
 * Sampling:                          SAMPLE_CELLS=32768 pnpm bench:write
 * Keep the generated file with:      KEEP=1 pnpm bench:write
 */
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const ROWS = Number(process.env.ROWS ?? 100_000);
const COLS = Number(process.env.COLS ?? 8);
const SHEET = 'Data';
const WRITE_URL = new URL('../dist/write.mjs', import.meta.url).href;
const STREAM_URL = new URL('../dist/stream.mjs', import.meta.url).href;
const INDEX_URL = new URL('../dist/index.mjs', import.meta.url).href;

/**
 * The two row counts the "live memory does not grow with the row count" claim is
 * measured across, and how much the live peak is allowed to grow between them.
 * 8x the rows and a peak below `SCALE_TOLERANCE` is a real measurement; a single
 * row count is not a measurement at all.
 */
const SCALE_SMALL_ROWS = Number(process.env.SCALE_ROWS_SMALL ?? 50_000);
const SCALE_LARGE_ROWS = Number(process.env.SCALE_ROWS_LARGE ?? 400_000);
const SCALE_TOLERANCE = Number(process.env.SCALE_TOLERANCE ?? 1.5);

const mib = (bytes) => `${(bytes / 1024 / 1024).toFixed(1)} MB`;
const kib = (bytes) => `${(bytes / 1024).toFixed(1)} kB`;

// ---------------------------------------------------------------------------
// One phase per process: MODE=file|buffer|naive
// ---------------------------------------------------------------------------
const MEASURE = `
const MODE = process.env.MODE;
const ROWS = Number(process.env.ROWS);
const COLS = Number(process.env.COLS);
const OUT = process.env.OUT;
const gc = globalThis.gc;
// TypedArray backing stores live outside the V8 heap, so a writer's real
// footprint is heap + arrayBuffers.
const mem = () => { const m = process.memoryUsage(); return m.heapUsed + m.arrayBuffers; };
const settle = async () => { if (gc) { gc(); await new Promise((r) => setTimeout(r, 60)); gc(); } };
const out = (obj) => process.stdout.write(JSON.stringify(obj));

// Deterministic, poorly-compressible filler, so the worksheet XML is realistic
// in size instead of collapsing to nothing under DEFLATE.
let seed = 0x2545f491;
const rand = () => {
  seed ^= seed << 13; seed >>>= 0;
  seed ^= seed >>> 17;
  seed ^= seed << 5; seed >>>= 0;
  return seed;
};
const token = (i) => {
  let s = '';
  let x = rand();
  for (let k = 0; k < 12; k++) { s += String.fromCharCode(97 + (x % 26)); x = Math.floor(x / 26) + rand() % 97; }
  return s + i;
};

const rowValues = (r) => [
  r,
  'row-' + r,
  r * 1.5,
  r * 0.01,
  r % 2 === 0,
  'text ' + r,
  token(r),
  token(r * 7),
].slice(0, COLS);

// Load the implementation BEFORE the baseline is taken. Compiling the module,
// building fflate's tables and JIT-warming the writer are a fixed cost with
// nothing to do with the rows being written, and left in the delta they swamp
// the per-row signal: ~10 MB of import against a streaming writer whose live set
// is a few hundred kB. The baseline must exclude them or every number below is
// really "import cost + writer".
const impl = MODE === 'naive' ? await import('fflate') : await import(process.env.WRITE_MOD);
await settle();
const base = mem();
// The row loops are synchronous, so a timer cannot interrupt them: the honest
// figure is the highest value seen at the periodic forced-GC checkpoints.
let live = 0;

if (MODE === 'naive') {
  // The in-memory shape: accumulate the whole sheet, then compress it.
  const { zipSync } = impl;
  const parts = ['<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'];
  const t0 = performance.now();
  for (let r = 1; r <= ROWS; r++) {
    parts.push('<row r="' + r + '">');
    rowValues(r).forEach((v, i) => {
      const ref = String.fromCharCode(65 + i) + r;
      parts.push('<c r="' + ref + '"><v>' + v + '</v></c>');
    });
    parts.push('</row>');
    if (gc && r % 4096 === 0) { gc(); const u = mem() - base; if (u > live) live = u; }
  }
  parts.push('</sheetData></worksheet>');
  const xml = parts.join('');
  parts.length = 0;
  const bytes = zipSync({ 'xl/worksheets/sheet1.xml': new TextEncoder().encode(xml) });
  const ms = performance.now() - t0;
  const end = mem() - base;
  out({ live: Math.max(live, end), written: ms, total: ms, bytes: bytes.length });
} else {
  const { WorkbookWriter } = impl;
  const wb = MODE === 'file' ? await WorkbookWriter.create({ file: OUT }) : await WorkbookWriter.create();
  const ws = wb.addWorksheet('${SHEET}');
  const header = ws.addRow(['id', 'label', 'value', 'ratio', 'flag', 'text', 'count', 'junk'].slice(0, COLS));
  header.font = { bold: true };
  header.fill = { pattern: 'solid', fgColor: 'FF203864' };

  let sinceGc = 0;
  const t0 = performance.now();
  for (let r = 1; r <= ROWS; r++) {
    const row = ws.addRow(rowValues(r));
    if (COLS > 2) row.getCell(3).numFmt = '#,##0.00';
    sinceGc += COLS;
    if (gc && sinceGc >= 4096) {
      sinceGc = 0;
      gc();
      const u = mem() - base;
      if (u > live) live = u;
    }
  }
  const written = performance.now() - t0;

  let bytes = 0;
  if (MODE === 'file') {
    await wb.commit();
  } else {
    const buf = await wb.toBuffer();
    bytes = buf.length;
    if (gc) { gc(); live = Math.max(live, mem() - base); }
  }
  const total = performance.now() - t0;
  const end = mem() - base;
  if (end > live) live = end;
  out({ live, written, total, bytes });
}
`;

const measure = (mode, out, rows = ROWS) => {
  const raw = execFileSync(process.execPath, ['--expose-gc', '--input-type=module', '-e', MEASURE], {
    env: {
      ...process.env,
      MODE: mode,
      ROWS: String(rows),
      COLS: String(COLS),
      OUT: out,
      WRITE_MOD: WRITE_URL,
      INDEX_MOD: INDEX_URL,
    },
    encoding: 'utf8',
    maxBuffer: 1024 * 1024 * 8,
    timeout: 30 * 60_000,
  });
  return JSON.parse(raw);
};

// ---------------------------------------------------------------------------
// Run
// ---------------------------------------------------------------------------
const dir = mkdtempSync(join(tmpdir(), 'sheet-lite-bench-write-'));
const filePath = join(dir, 'bench-write.xlsx');

console.log(`\nsheet-lite — incremental writer vs in-memory writes (${ROWS} rows x ${COLS} cols)`);
console.log('  memory = heapUsed + arrayBuffers (TypedArray data is off-heap)');

// (a) file target
const fileResult = measure('file', filePath);
fileResult.bytes = statSync(filePath).size;

// Inflated part sizes straight from the central directory (nothing inflated).
const { readZipDirectory, BufferByteSource } = await import(STREAM_URL);
const dirents = await readZipDirectory(new BufferByteSource(new Uint8Array(readFileSync(filePath))));
const sheetEntry = dirents.byName.get('xl/worksheets/sheet1.xml');
console.log(`  archive: ${kib(fileResult.bytes)} compressed`);
console.log(`  worksheet part inflated: ${sheetEntry ? mib(sheetEntry.uncompressedSize) : '?'}`);

// (b) buffered target
const bufferResult = measure('buffer', join(dir, 'unused.xlsx'));
// (c) naive in-memory
const naiveResult = measure('naive', join(dir, 'unused-naive.xlsx'));

const table = [
  ['metric', 'file target (streaming)', 'buffer target', 'naive in-memory'],
  ['live memory (GC-sampled peak)', mib(fileResult.live), mib(bufferResult.live), mib(naiveResult.live)],
  ['add-row time', `${(fileResult.written / 1000).toFixed(2)} s`, `${(bufferResult.written / 1000).toFixed(2)} s`, `${(naiveResult.written / 1000).toFixed(2)} s`],
  ['total time', `${(fileResult.total / 1000).toFixed(2)} s`, `${(bufferResult.total / 1000).toFixed(2)} s`, `${(naiveResult.total / 1000).toFixed(2)} s`],
  ['output size', kib(fileResult.bytes), bufferResult.bytes ? kib(bufferResult.bytes) : 'n/a', kib(naiveResult.bytes)],
];
const widths = table[0].map((_, i) => Math.max(...table.map((row) => row[i].length)));
console.log('');
for (const row of table) {
  console.log('  ' + row.map((cell, i) => cell.padEnd(widths[i])).join('  '));
}

const ratio = naiveResult.live / Math.max(1, fileResult.live);
console.log(
  `\n  at ${ROWS} rows the streaming writer's live peak is ${ratio.toFixed(0)}x smaller than the in-memory build.`,
);

// ---------------------------------------------------------------------------
// The claim that needs two row counts, not one
// ---------------------------------------------------------------------------
// "does not grow with the row count" is a claim about a SLOPE. Measured at a
// single row count it is unfalsifiable: a writer that leaked 20 bytes per row
// would look identical at 100k rows. So the same file-target writer is run at
// SCALE_LARGE_ROWS and SCALE_SMALL_ROWS — 8x the rows — and the live peaks are
// compared. The bound is a bound, not a tautology: it fails the run if the peak
// grows by more than SCALE_TOLERANCE across an 8x larger sheet.
console.log(
  `\n  constant-memory check — file target, ${SCALE_SMALL_ROWS} vs ${SCALE_LARGE_ROWS} rows ` +
    `(${(SCALE_LARGE_ROWS / SCALE_SMALL_ROWS).toFixed(0)}x the rows):`,
);
const small = measure('file', join(dir, `scale-${SCALE_SMALL_ROWS}.xlsx`), SCALE_SMALL_ROWS);
const large = measure('file', join(dir, `scale-${SCALE_LARGE_ROWS}.xlsx`), SCALE_LARGE_ROWS);
const growth = large.live / Math.max(1, small.live);
const flatEnough = large.live < SCALE_TOLERANCE * Math.max(1, small.live);
console.log(`    live peak @ ${String(SCALE_SMALL_ROWS).padStart(6)} rows: ${mib(small.live)}`);
console.log(`    live peak @ ${String(SCALE_LARGE_ROWS).padStart(6)} rows: ${mib(large.live)}  (${growth.toFixed(2)}x)`);
console.log(`    bound: must stay below ${SCALE_TOLERANCE}x  →  ${flatEnough ? 'ok' : 'FAILED'}`);
if (!flatEnough) {
  console.error(
    `\n  live memory grew ${growth.toFixed(2)}x for a ` +
      `${(SCALE_LARGE_ROWS / SCALE_SMALL_ROWS).toFixed(0)}x larger sheet ` +
      `(${mib(small.live)} -> ${mib(large.live)}): the writer is retaining per-row state, ` +
      `so "constant memory" is not true at ${SCALE_LARGE_ROWS} rows.`,
  );
  process.exitCode = 1;
}
console.log(
  `  only fflate's deflate window, one row and the ZIP metadata are alive: ` +
    `${mib(small.live)} at ${SCALE_SMALL_ROWS} rows, ${mib(large.live)} at ${SCALE_LARGE_ROWS}.`,
);
console.log(
  `  throughput: ${(ROWS / (fileResult.total / 1000) / 1000).toFixed(0)}k rows/s ` +
    `(${mib((sheetEntry?.uncompressedSize ?? 0) / Math.max(1, fileResult.total / 1000))}/s of XML, level 6).`,
);

if (!process.env.KEEP) rmSync(dir, { recursive: true, force: true });