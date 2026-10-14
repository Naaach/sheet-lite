/**
 * Real-world corpus validator.
 *
 * Opens every `.xlsx` / `.xlsm` in `test/fixtures/real/`, exercises read +
 * no-op round-trip + a single-cell edit, and (when LibreOffice is installed)
 * checks that the edited file re-opens without a repair prompt.
 *
 * Run after a build:   pnpm build && node scripts/validate-corpus.mjs
 *
 * Exit code is non-zero only when a file that OPENED also broke an invariant.
 * Unsupported files (zip64 / encrypted / data descriptors) are reported as
 * "unsupported" and do NOT fail the run as long as the rejection is coded.
 */
import { existsSync, readdirSync, readFileSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { Workbook, SheetLiteError } from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const DIR = join(here, '..', 'test', 'fixtures', 'real');
const files = existsSync(DIR)
  ? readdirSync(DIR)
      .filter((name) => /\.(xlsx|xlsm)$/i.test(name))
      .sort()
  : [];

const kib = (bytes) => `${(bytes / 1024).toFixed(1)} kB`;
const equal = (a, b) => Buffer.from(a).equals(Buffer.from(b));

function findLibreOffice() {
  for (const bin of ['soffice', 'libreoffice']) {
    const probe = spawnSync(bin, ['--version'], { stdio: 'ignore', shell: false });
    if (!probe.error) return bin;
  }
  return null;
}

function libreOfficeOpens(bin, bytes, workdir) {
  // Write the buffer, then ask LibreOffice to convert it. A corrupt file makes
  // the conversion fail (non-zero exit or no output).
  const input = join(workdir, 'probe.xlsx');
  writeFileSync(input, bytes);
  const res = spawnSync(bin, ['--headless', '--convert-to', 'xlsx', '--outdir', workdir, input], {
    stdio: 'ignore',
  });
  return !res.error && res.status === 0;
}

if (files.length === 0) {
  console.log('No real-world files yet. Drop .xlsx/.xlsm into test/fixtures/real/ and re-run.');
  process.exit(0);
}

const soffice = findLibreOffice();
console.log(`Corpus: ${files.length} file(s) · LibreOffice gate: ${soffice ?? 'not installed (skipped)'}\n`);

let failures = 0;
const rows = [];

for (const file of files) {
  const bytes = new Uint8Array(readFileSync(join(DIR, file)));
  const result = { file, size: kib(bytes.length), status: '', detail: '' };

  let wb;
  try {
    wb = Workbook.read(bytes);
  } catch (error) {
    if (error instanceof SheetLiteError) {
      result.status = 'unsupported';
      result.detail = error.code;
      rows.push(result);
      continue;
    }
    result.status = 'FAIL';
    result.detail = String(error);
    failures += 1;
    rows.push(result);
    continue;
  }

  try {
    let cellCount = 0;
    for (const ws of wb.worksheets) ws.eachRow((row) => row.eachCell(() => (cellCount += 1)));

    const noOpOk = equal(wb.toBuffer(), bytes);

    const edited = Workbook.read(bytes);
    const target = edited.worksheets[0];
    target.getCell('ZZ1').value = 'sheet-lite corpus probe';
    const out = edited.toBuffer();
    const before = Workbook.read(bytes).parts();
    const after = Workbook.read(out).parts();
    let onlyTargetChanged = true;
    for (const name of before.names()) {
      if (name === target.partName) continue;
      if (!equal(before.get(name), after.get(name))) onlyTargetChanged = false;
    }

    let loOk = null;
    if (soffice) {
      const workdir = mkdtempSync(join(tmpdir(), 'sheet-lite-corpus-'));
      try {
        loOk = libreOfficeOpens(soffice, out, workdir);
      } finally {
        rmSync(workdir, { recursive: true, force: true });
      }
    }

    const ok = noOpOk && onlyTargetChanged && (loOk === null || loOk);
    result.status = ok ? 'ok' : 'FAIL';
    result.detail = [
      `${wb.sheetNames.length} sheet(s)`,
      `${cellCount} cell(s)`,
      `no-op ${noOpOk ? '✓' : '✗'}`,
      `edit ${onlyTargetChanged ? '✓' : '✗'}`,
      soffice ? `libreoffice ${loOk ? '✓' : '✗'}` : 'libreoffice n/a',
    ].join(' · ');
    if (!ok) failures += 1;
  } catch (error) {
    result.status = 'FAIL';
    result.detail = String(error);
    failures += 1;
  }
  rows.push(result);
}

const width = Math.max(...rows.map((r) => r.file.length), 4);
for (const r of rows) {
  console.log(`${r.file.padEnd(width)}  ${r.size.padStart(9)}  ${r.status.padEnd(11)}  ${r.detail}`);
}

console.log(`\n${rows.filter((r) => r.status === 'ok').length} ok · ${rows.filter((r) => r.status === 'unsupported').length} unsupported · ${failures} failed`);
process.exit(failures > 0 ? 1 : 0);
