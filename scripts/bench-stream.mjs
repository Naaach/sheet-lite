/**
 * Streaming benchmark: eager (whole-archive) read vs `sheet-lite/stream`.
 *
 * Run with:  pnpm bench:stream   (builds dist first, then `node --expose-gc`)
 *
 * The fixture (~100k rows x 8 cols, ~30 MB of worksheet XML) is generated with
 * ExcelJS, and each phase is measured in its OWN fresh node process, so nothing
 * pollutes the numbers:
 *   (a) eager   `Workbook.read(file)` — inflates + parses the whole package
 *   (b) stream  `StreamingWorkbook.open(path)` + a full `readRows` walk
 *
 * Memory is sampled from `process.memoryUsage().heapUsed` every 2 ms during
 * each phase. For streaming we also report a *live* peak: the same walk with
 * periodic forced GC, so the number reflects memory the reader actually holds
 * rather than short-lived per-row garbage.
 *
 * NOTE: eager `Worksheet.eachRow()` is quadratic today (it re-scans the cell
 * map for every row) and is therefore NOT part of the eager timing — the eager
 * cost measured here is `Workbook.read`, which is also where the memory goes.
 *
 * Override the size with env vars:  ROWS=50000 COLS=8 pnpm bench:stream
 * Keep the generated file with:     KEEP=1 pnpm bench:stream
 */
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const ROWS = Number(process.env.ROWS ?? 100_000);
const COLS = Number(process.env.COLS ?? 8);
const TARGET_SHEET = 'Data';
const INDEX_URL = new URL('../dist/index.mjs', import.meta.url).href;
const STREAM_URL = new URL('../dist/stream.mjs', import.meta.url).href;

const mib = (bytes) => `${(bytes / 1024 / 1024).toFixed(1)} MB`;
const kib = (bytes) => `${(bytes / 1024).toFixed(1)} kB`;

// ---------------------------------------------------------------------------
// Fixture generation (child process, so ExcelJS allocations stay out of the way)
// ---------------------------------------------------------------------------
const GENERATOR = `
const ExcelJS = require(process.env.EXCELJS);
const ROWS = Number(process.env.ROWS);
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('${TARGET_SHEET}');
for (let r = 1; r <= ROWS; r += 1) {
  ws.getCell(r, 1).value = 'row-' + r;
  ws.getCell(r, 2).value = r;
  ws.getCell(r, 3).value = r * 1.5;
  ws.getCell(r, 4).value = new Date(Date.UTC(2020, 0, 1) + r * 86400000);
  ws.getCell(r, 4).numFmt = 'yyyy-mm-dd';
  ws.getCell(r, 5).value = r % 2 === 0;
  ws.getCell(r, 6).value = r * 0.01;
  ws.getCell(r, 7).value = 'text ' + r;
  ws.getCell(r, 8).value = r;
}
ws.getCell(2, 2).value = { formula: 'B1+B3', result: 4 };
wb.xlsx.writeFile(process.env.OUT).then(() => process.stdout.write('ok'));
`;

// ---------------------------------------------------------------------------
// Measurement child: MODE=eager|stream, one phase, fresh process, prints JSON
// ---------------------------------------------------------------------------
const MEASURE = `
const MODE = process.env.MODE;
const FILE = process.env.FILE;
const SHEET = '${TARGET_SHEET}';
const gc = globalThis.gc;
const heap = () => process.memoryUsage().heapUsed;
const settle = async () => { if (gc) { gc(); await new Promise((r) => setTimeout(r, 60)); gc(); } };
const startTimer = (base) => {
  let peak = heap() - base;
  const iv = setInterval(() => { const u = heap() - base; if (u > peak) peak = u; }, 2);
  return { stop: () => { clearInterval(iv); return Math.max(peak, heap() - base); } };
};
const out = (obj) => process.stdout.write(JSON.stringify(obj));

await settle();
const base = heap();

if (MODE === 'eager') {
  const { Workbook } = await import(process.env.INDEX_MOD);
  const fs = await import('node:fs');
  const timer = startTimer(base);
  const t0 = performance.now();
  let held = Workbook.read(fs.readFileSync(FILE));
  void held.getWorksheet(SHEET).getCell('A1').value;
  const ms = performance.now() - t0;
  const peak = timer.stop();
  const retained = heap() - base; // still holding the workbook
  void held;
  out({ peak, retained, ms });
} else {
  const { StreamingWorkbook } = await import(process.env.STREAM_MOD);
  const wb = await StreamingWorkbook.open(FILE);
  let cells = 0;
  let live = 0;
  let sinceGc = 0;
  const timer = startTimer(base);
  const t0 = performance.now();
  for await (const row of wb.readRows(SHEET)) {
    cells += row.cells.length;
    if (gc) {
      sinceGc += row.cells.length;
      if (sinceGc >= 4096) {
        sinceGc = 0;
        gc();
        const used = heap() - base;
        if (used > live) live = used;
      }
    }
  }
  const ms = performance.now() - t0;
  const peak = timer.stop();
  if (gc) { gc(); live = Math.max(live, heap() - base); }
  await wb.close();
  out({ peak, live, ms, cells });
}
`;

const measure = (mode, file, indexMod, streamMod) => {
  const raw = execFileSync(process.execPath, ['--expose-gc', '--input-type=module', '-e', MEASURE], {
    env: { ...process.env, MODE: mode, FILE: file, INDEX_MOD: indexMod, STREAM_MOD: streamMod },
    encoding: 'utf8',
    timeout: 10 * 60_000,
  });
  return JSON.parse(raw);
};

// ---------------------------------------------------------------------------
// Run
// ---------------------------------------------------------------------------
const dir = mkdtempSync(join(tmpdir(), 'sheet-lite-bench-'));
const filePath = join(dir, 'bench.xlsx');
writeFileSync(join(dir, 'generate.cjs'), GENERATOR);
execFileSync(process.execPath, [join(dir, 'generate.cjs')], {
  env: {
    ...process.env,
    ROWS: String(ROWS),
    OUT: filePath,
    EXCELJS: createRequire(import.meta.url).resolve('exceljs'),
  },
  stdio: 'ignore',
  timeout: 10 * 60_000,
});

const { size } = await import('node:fs/promises').then((fs) => fs.stat(filePath));
console.log(`\nsheet-lite — eager vs streaming (${ROWS} rows x ${COLS} cols)`);
console.log(`  archive: ${kib(size)} compressed (${filePath})`);

// Inflated part sizes from the central directory (nothing is inflated here).
const { readZipDirectory, BufferByteSource } = await import(STREAM_URL);
const dirents = await readZipDirectory(new BufferByteSource(readFileSync(filePath)));
const sheetEntry = dirents.byName.get('xl/worksheets/sheet1.xml');
const sharedEntry = dirents.byName.get('xl/sharedStrings.xml');
console.log(`  worksheet part inflated: ${sheetEntry ? mib(sheetEntry.uncompressedSize) : '?'}`);
console.log(`  sharedStrings inflated: ${sharedEntry ? mib(sharedEntry.uncompressedSize) : '0 kB'}`);

const eager = measure('eager', filePath, INDEX_URL, STREAM_URL);
const stream = measure('stream', filePath, INDEX_URL, STREAM_URL);

const table = [
  ['metric', 'eager (Workbook.read)', 'stream (StreamingWorkbook)'],
  ['peak heap (sampled)', mib(eager.peak), mib(stream.peak)],
  ['live heap (after GC)', `${mib(eager.retained)} (retained)`, mib(stream.live)],
  ['read / walk time', `${(eager.ms / 1000).toFixed(2)} s`, `${(stream.ms / 1000).toFixed(2)} s`],
];
const widths = table[0].map((_, i) => Math.max(...table.map((r) => r[i].length)));
console.log('');
for (const row of table) {
  console.log('  ' + row.map((cell, i) => cell.padEnd(widths[i])).join('  '));
}
console.log(`  streamed cells: ${stream.cells.toLocaleString()}`);
console.log(
  `\n  streaming live heap is ${(eager.retained / Math.max(1, stream.live)).toFixed(0)}x smaller than the eager retained heap; ` +
    `sharedStrings (${sharedEntry ? mib(sharedEntry.uncompressedSize) : '0'}) is the only part read whole.`,
);

if (!process.env.KEEP) rmSync(dir, { recursive: true, force: true });