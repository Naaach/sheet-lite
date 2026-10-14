/**
 * Benchmark: sheet-lite vs ExcelJS vs SheetJS.
 * Run with: pnpm bench   (builds dist first, then executes this file)
 *
 * Sizes are the published, gzipped, whole-package figures (minified, including
 * dependencies) as measured by bundlephobia for the versions in devDependencies.
 */
import { gzipSync } from 'node:zlib';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { Workbook, readZip, writeZip } from '../dist/index.mjs';

const ROWS = 5000;
const COLS = 8;

const PUBLISHED_GZIP = {
  'sheet-lite': 8.7,
  fflate: 11.8,
  ExcelJS: 249.9,
  SheetJS: 136.4,
};

async function makeFile() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  for (let r = 1; r <= ROWS; r += 1) {
    ws.getCell(r, 1).value = `row-${r}`;
    ws.getCell(r, 2).value = r;
    ws.getCell(r, 3).value = r * 1.5;
    ws.getCell(r, 4).value = new Date(Date.UTC(2020, 0, 1) + r * 86_400_000);
    ws.getCell(r, 5).value = r % 2 === 0;
    ws.getCell(r, 6).value = r * 0.01;
    ws.getCell(r, 7).value = `text ${r}`;
    ws.getCell(r, 8).value = r;
  }
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function time(label, runs, fn) {
  await fn();
  const samples = [];
  for (let i = 0; i < runs; i += 1) {
    const t0 = performance.now();
    await fn();
    samples.push(performance.now() - t0);
  }
  samples.sort((a, b) => a - b);
  return { label, median: samples[Math.floor(samples.length / 2)] };
}

const file = await makeFile();
console.log(`\nGenerated file: ${(file.length / 1024).toFixed(1)} kB · ${ROWS} rows x ${COLS} cols\n`);

const results = [
  await time('sheet-lite', 5, () => Workbook.read(file).getWorksheet('Data')?.rowCount),
  await time('SheetJS', 5, () => XLSX.read(file, { type: 'buffer', cellDates: true }).SheetNames.length),
  await time('ExcelJS', 5, async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(file);
  }),
];

console.log('Read time (median of 5):');
for (const r of results) console.log(`  ${r.label.padEnd(12)} ${r.median.toFixed(1)} ms`);

// WRITE benchmark: open the generated file, edit one cell, serialize.
// Untouched entries are carried over verbatim (raw deflate reuse).
const writeResults = [
  await time('sheet-lite write', 5, () => {
    const wb = Workbook.read(file);
    wb.getWorksheet('Data').getCell('B2').value = 987654.321;
    return wb.toBuffer();
  }),
  await time('ExcelJS write', 5, async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(file);
    wb.getWorksheet('Data').getCell('B2').value = 987654.321;
    return Buffer.from(await wb.xlsx.writeBuffer());
  }),
];

console.log('\nWrite time (read + edit one cell + serialize, median of 5):');
for (const r of writeResults) console.log(`  ${r.label.padEnd(12)} ${r.median.toFixed(1)} ms`);

// Serialize-only (workbooks pre-built, so read/parse cost is excluded).
// The dirty sheet is ~85% of this file, so re-compressing it dominates the
// first case; editing a SMALL part is where verbatim reuse shines.
const editedBig = Workbook.read(file);
editedBig.getWorksheet('Data').getCell('B2').value = 987654.321;
const editedSmall = Workbook.read(file);
editedSmall.parts().set(
  'docProps/core.xml',
  new TextEncoder().encode(new TextDecoder().decode(editedSmall.parts().get('docProps/core.xml')).replace('sheet-lite', 'sheet-lite!')),
);
const rawArchive = readZip(file);

const zipOnly = [
  await time('serialize: one cell edited', 7, () => editedBig.toBuffer()),
  await time('serialize: small part edited', 7, () => editedSmall.toBuffer()),
  await time('serialize: raw pass-through', 7, () => writeZip(rawArchive)),
];

console.log('\nSerialize only (toBuffer on a pre-built workbook, median of 7):');
for (const r of zipOnly) console.log(`  ${r.label.padEnd(28)} ${r.median.toFixed(1)} ms`);

const distGzip = gzipSync(readFileSync(fileURLToPath(new URL('../dist/index.mjs', import.meta.url)))).length / 1024;
console.log('\nPublished size (gzipped, whole package incl. deps):');
console.log(`  sheet-lite   ${distGzip.toFixed(1)} kB  +  fflate ${PUBLISHED_GZIP.fflate} kB  =  ~${(distGzip + PUBLISHED_GZIP.fflate).toFixed(1)} kB   (0 transitive deps)`);
console.log(`  ExcelJS      ${PUBLISHED_GZIP.ExcelJS} kB   (9 deps)`);
console.log(`  SheetJS      ${PUBLISHED_GZIP.SheetJS} kB   (7 deps)`);
