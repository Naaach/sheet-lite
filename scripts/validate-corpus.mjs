/**
 * Real-world corpus validator.
 *
 * Opens every `.xlsx` / `.xlsm` in `test/fixtures/real/`, exercises read +
 * no-op round-trip + a single-cell edit, and (when LibreOffice is installed)
 * checks that the edited file re-opens without a repair prompt.
 *
 * The legacy binaries in the same folder (`.xlsb`, `.xls`) are classified
 * explicitly as UNSUPPORTED: they must be refused, with the exact coded error
 * sheet-lite promises. Accepting one of them is a failure (the format support
 * claim changed without the corpus noticing), as is refusing it with a
 * different code or with a raw crash.
 *
 * Run after a build:   pnpm build && node scripts/validate-corpus.mjs
 *
 * Exit code is non-zero when a file that OPENED broke an invariant, when a
 * file was refused for an unexpected reason, or when an unsupported format
 * started being accepted.
 */
import { existsSync, readdirSync, readFileSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { Workbook, SheetLiteError } from '../dist/index.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const DIR = join(here, '..', 'test', 'fixtures', 'real');
/** Formats sheet-lite claims to read. */
const SUPPORTED_EXT = /\.(xlsx|xlsm)$/i;
/**
 * Formats that are deliberately NOT supported. They must keep being refused
 * with these exact codes (keep in sync with `test/corpus.test.ts`):
 * `NOT_A_ZIP` is what the low-level ZIP layer reports for "not an archive".
 */
const UNSUPPORTED = {
  xlsb: { code: 'NOT_XLSX', match: /binary Excel workbook \(\.xlsb\)/ },
  xls: { code: 'NOT_XLSX', match: /OLE2 compound file/ },
};
const UNSUPPORTED_EXT = /\.(xlsb|xls)$/i;

const entries = existsSync(DIR) ? readdirSync(DIR).filter((name) => !/\.(md|txt)$/i.test(name)).sort() : [];
const files = entries.filter((name) => SUPPORTED_EXT.test(name));
const binaries = entries.filter((name) => UNSUPPORTED_EXT.test(name));
// Nothing may fall through both filters unnoticed (a renamed .xls, a .csv, ...).
const unclassified = entries.filter((name) => !SUPPORTED_EXT.test(name) && !UNSUPPORTED_EXT.test(name));

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

if (files.length === 0 && binaries.length === 0) {
  console.log('No real-world files yet. Drop .xlsx/.xlsm into test/fixtures/real/ and re-run.');
  process.exit(0);
}

const soffice = findLibreOffice();
console.log(
  `Corpus: ${files.length} readable file(s) · ${binaries.length} unsupported binary file(s) · ` +
    `LibreOffice gate: ${soffice ?? 'not installed (skipped)'}\n`,
);

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

// Legacy binaries (.xlsb / .xls): sheet-lite does not read them, so the only
// correct outcome is an honest, coded refusal. Anything else — a silent
// accept, a wrong code, a message that does not name the format, or a crash —
// fails the run: the documented rejection is part of the contract.
const legacyRows = [];
for (const file of binaries) {
  const bytes = new Uint8Array(readFileSync(join(DIR, file)));
  const ext = file.match(UNSUPPORTED_EXT)[0].slice(1).toLowerCase();
  const expected = UNSUPPORTED[ext];
  const row = { file, size: kib(bytes.length), status: 'unsupported', detail: '' };
  legacyRows.push(row);

  if (!expected) {
    row.status = 'FAIL';
    row.detail = `no expected rejection declared for .${ext} — extend UNSUPPORTED`;
  } else {
    try {
      Workbook.read(bytes);
      row.status = 'FAIL';
      row.detail = `.${ext} is not supported but was accepted — refusing it is the contract`;
    } catch (caught) {
      if (!(caught instanceof SheetLiteError)) {
        row.status = 'FAIL';
        row.detail = `crashed instead of failing cleanly: ${String(caught)}`;
      } else if (caught.code !== expected.code) {
        row.status = 'FAIL';
        row.detail = `${caught.code} — expected ${expected.code}`;
      } else if (!expected.match.test(caught.message)) {
        row.status = 'FAIL';
        row.detail = `${caught.code}: "${caught.message}" — does not name the format`;
      } else {
        row.detail = `${caught.code} · refused as expected`;
      }
    }
  }
  if (row.status === 'FAIL') failures += 1;
}

// A file that is neither readable nor a known legacy binary is a silent gap in
// the corpus: report it instead of ignoring it.
for (const file of unclassified) {
  legacyRows.push({
    file,
    size: '-',
    status: 'FAIL',
    detail: 'unclassified corpus file (neither a supported .xlsx/.xlsm nor a known legacy binary)',
  });
  failures += 1;
}

const width = Math.max(...[...rows, ...legacyRows].map((r) => r.file.length), 4);
const print = (list, heading) => {
  if (list.length === 0) return;
  if (heading) console.log(`${heading}\n`);
  for (const r of list) {
    console.log(`${r.file.padEnd(width)}  ${String(r.size).padStart(9)}  ${r.status.padEnd(11)}  ${r.detail}`);
  }
  console.log('');
};
print(rows, 'Supported (.xlsx/.xlsm):');
print(legacyRows, 'Unsupported by design (.xlsb/.xls) — must stay refused:');

const okCount = rows.filter((r) => r.status === 'ok').length;
const unsupportedCount = rows.filter((r) => r.status === 'unsupported').length + legacyRows.filter((r) => r.status === 'unsupported').length;
console.log(`${okCount} ok · ${unsupportedCount} unsupported · ${failures} failed`);
process.exit(failures > 0 ? 1 : 0);
