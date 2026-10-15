import { describe, it, expect } from 'vitest';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { Workbook, SheetLiteError } from '../src/index';

const here = dirname(fileURLToPath(import.meta.url));
const DIR = join(here, 'fixtures', 'real');

/** The fixture folder also carries its own README; that is not a test input. */
const isDoc = (name: string): boolean => /\.(md|txt)$/i.test(name);
const ALL = existsSync(DIR) ? readdirSync(DIR).filter((name) => !isDoc(name)).sort() : [];

/** The read corpus: the OOXML XML formats sheet-lite claims to support. */
const FILES = ALL.filter((name) => /\.(xlsx|xlsm)$/i.test(name));

/**
 * Everything else that is a spreadsheet file but NOT an OOXML XML package:
 * `.xlsb` (binary workbook inside a ZIP) and `.xls` (legacy OLE/CFB compound
 * file). `FILES` above must exclude all of them — see the "rejection" suite
 * below, which asserts each one is refused with a coded error.
 */
const UNSUPPORTED_EXT = /\.(xlsb|xls)$/i;
const UNSUPPORTED = ALL.filter((name) => UNSUPPORTED_EXT.test(name));

/**
 * `.xlsx`/`.xlsm` corpus files sheet-lite currently refuses to open, mapped to
 * the exact error code it refuses them with.
 *
 * This list is intentionally EMPTY. An empty list means the strong, honest
 * assertion: every real-world `.xlsx`/`.xlsm` in the corpus opens. A file may
 * only be added here deliberately, together with the reason it cannot be
 * opened — the suite below then fails loudly if the file starts opening (so
 * the entry cannot rot) and if the code ever changes.
 */
const KNOWN_UNSUPPORTED: Record<string, string> = {};

/** Number of non-XML binaries the fixture folder is expected to carry. */
const EXPECTED_UNSUPPORTED_COUNT = 14;

const equal = (a: Uint8Array, b: Uint8Array): boolean => Buffer.from(a).equals(Buffer.from(b));

const describeError = (error: unknown): string => {
  if (error instanceof SheetLiteError) return `${error.code}: ${error.message}`;
  return `${(error as Error)?.constructor?.name ?? typeof error}: ${String(error)}`;
};

/**
 * What the library reports TODAY for a non-OOXML spreadsheet, keyed by
 * extension. Both are `NOT_XLSX`: a wrong format is not an attack, so it must
 * never surface as `ZIP_SECURITY` (which stays reserved for genuinely hostile
 * archives — zip bombs, traversal, limit violations).
 *
 * Do NOT "fix" these strings here without changing `src/`: this table is the
 * observation, and the exact code + message are asserted below.
 */
const OBSERVED_REJECTION: Record<string, { code: string; message: string }> = {
  // `.xlsb` IS a ZIP, so it gets past the ZIP layer and is rejected by the
  // package reader, which now recognises the binary workbook part
  // (`xl/workbook.bin`) and names the real cause.
  xlsb: {
    code: 'NOT_XLSX',
    message: 'Not an XLSX package: binary Excel workbook (.xlsb); sheet-lite reads .xlsx/.xlsm only',
  },
  // `.xls` is a legacy OLE/CFB compound file, NOT a ZIP at all: it is sniffed
  // before the ZIP layer runs, so it cannot be mistaken for a corrupt or
  // malicious archive. (Encrypted OOXML packages use the same container.)
  xls: {
    code: 'NOT_XLSX',
    message:
      'Not an XLSX package: OLE2 compound file (legacy .xls or encrypted OOXML); sheet-lite reads .xlsx/.xlsm only',
  },
};

describe('real-world corpus', () => {
  it('splits the fixture folder into the read corpus and the refused binaries', () => {
    // The read filter must never pick up a format the library does not support.
    expect(FILES.filter((name) => UNSUPPORTED_EXT.test(name))).toEqual([]);
    // ...and nothing may fall through both filters unnoticed (e.g. a new .csv,
    // .ods or .xlsm renamed): every spreadsheet-ish file must be classified.
    const classified = new Set([...FILES, ...UNSUPPORTED]);
    expect(ALL.filter((name) => !classified.has(name))).toEqual([]);
    expect(UNSUPPORTED.length).toBe(EXPECTED_UNSUPPORTED_COUNT);
  });

  it('has no stale entry in KNOWN_UNSUPPORTED', () => {
    for (const [file, code] of Object.entries(KNOWN_UNSUPPORTED)) {
      expect(FILES, `KNOWN_UNSUPPORTED names a file that is not in the corpus: ${file}`).toContain(file);
      expect(typeof code, `KNOWN_UNSUPPORTED entry without a code: ${file}`).toBe('string');
    }
  });

  if (FILES.length === 0) {
    it('is empty — drop real .xlsx files into test/fixtures/real/', () => {
      expect(FILES).toEqual([]);
    });
    return;
  }

  for (const file of FILES) {
    it(`${file}: opens, reads and round-trips without loss`, () => {
      const bytes = new Uint8Array(readFileSync(join(DIR, file)));

      let wb;
      try {
        wb = Workbook.read(bytes);
      } catch (error) {
        // An unexpected failure must NEVER pass unnoticed. The file either is
        // a declared known-unsupported one (and then the error must still be
        // the coded, clean rejection it was declared with), or this is a
        // regression and the test fails here, loudly, with the real error.
        const declared = KNOWN_UNSUPPORTED[file];
        expect(
          declared,
          `${file} failed to open but is not in KNOWN_UNSUPPORTED — add it deliberately or fix the regression. Observed: ${describeError(error)}`,
        ).toBeDefined();
        expect(error, `unexpected error: ${String(error)}`).toBeInstanceOf(SheetLiteError);
        expect((error as SheetLiteError).code).toBe(declared);
        return;
      }

      // A file listed as known-unsupported that now opens is just as much a
      // regression as the reverse: the list must not rot.
      expect(
        KNOWN_UNSUPPORTED[file],
        `${file} is listed in KNOWN_UNSUPPORTED but opened fine — remove the entry`,
      ).toBeUndefined();

      // Every sheet/row/cell is readable without throwing.
      expect(wb.sheetNames.length).toBeGreaterThan(0);
      for (const ws of wb.worksheets) {
        ws.eachRow((row) => row.eachCell((cell) => void cell.value));
      }

      // No-op write returns the original bytes verbatim.
      expect(equal(wb.toBuffer(), bytes)).toBe(true);

      // A single-cell edit leaves every other part byte-for-byte identical.
      const edited = Workbook.read(bytes);
      const target = edited.worksheets[0]!;
      const probe = 'sheet-lite corpus probe';
      target.getCell('ZZ1').value = probe;
      const out = edited.toBuffer();

      const before = Workbook.read(bytes).parts();
      const after = Workbook.read(out).parts();
      for (const name of before.names()) {
        if (name === target.partName) continue;
        expect(equal(before.get(name)!, after.get(name)!), `part changed: ${name}`).toBe(true);
      }

      // The probe reads back.
      expect(Workbook.read(out).getWorksheet(target.name)!.getCell('ZZ1').value).toBe(probe);
    });
  }
});

describe('unsupported spreadsheet formats are rejected clearly', () => {
  if (UNSUPPORTED.length === 0) {
    it('has no .xlsb/.xls binaries in the corpus', () => {
      expect(UNSUPPORTED).toEqual([]);
    });
    return;
  }

  for (const file of UNSUPPORTED) {
    const ext = (file.match(UNSUPPORTED_EXT)![0] as string).slice(1).toLowerCase();
    const expected = OBSERVED_REJECTION[ext]!;

    it(`${file}: refused as ${expected.code} — "${expected.message}"`, () => {
      const bytes = new Uint8Array(readFileSync(join(DIR, file)));

      let thrown: unknown;
      try {
        Workbook.read(bytes);
      } catch (error) {
        thrown = error;
      }

      // Clean, coded, specific rejection — never a crash, never a partially
      // built workbook, never silent acceptance.
      expect(thrown, `${file} was accepted: .${ext} is not supported`).toBeInstanceOf(SheetLiteError);
      const error = thrown as SheetLiteError;
      expect(error.code).toBe(expected.code);
      // The exact message is asserted so a change in what users are told shows
      // up here — and is reviewed on purpose.
      expect(error.message).toBe(expected.message);
      // A rejection is specific: it names something the user can act on.
      expect(error.message.length).toBeGreaterThan(0);
      expect(error).not.toBeInstanceOf(TypeError);
    });
  }

  it('reports every non-OOXML binary with one of the observed rejections', () => {
    const seen = new Set<string>();
    for (const file of UNSUPPORTED) {
      try {
        Workbook.read(new Uint8Array(readFileSync(join(DIR, file))));
        expect.unreachable(`${file} was accepted`);
      } catch (error) {
        expect(error, `${file} threw a non-SheetLiteError: ${String(error)}`).toBeInstanceOf(SheetLiteError);
        seen.add((error as SheetLiteError).code);
      }
    }
    // Two distinct messages, one code: a wrong format is never a security error.
    expect([...seen].sort()).toEqual(
      [...new Set(Object.values(OBSERVED_REJECTION).map((entry) => entry.code))].sort(),
    );
  });

  it('never reports a wrong-format file as a security problem', () => {
    for (const file of UNSUPPORTED) {
      try {
        Workbook.read(new Uint8Array(readFileSync(join(DIR, file))));
        expect.unreachable(`${file} was accepted`);
      } catch (error) {
        expect(
          (error as SheetLiteError).code,
          `${file} was reported as a security problem — ZIP_SECURITY is reserved for hostile archives`,
        ).not.toBe('ZIP_SECURITY');
      }
    }
  });
});