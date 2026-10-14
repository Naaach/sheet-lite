import { describe, it, expect } from 'vitest';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { Workbook, SheetLiteError } from '../src/index';

const here = dirname(fileURLToPath(import.meta.url));
const DIR = join(here, 'fixtures', 'real');
const FILES = existsSync(DIR)
  ? readdirSync(DIR)
      .filter((name) => /\.(xlsx|xlsm)$/i.test(name))
      .sort()
  : [];

const equal = (a: Uint8Array, b: Uint8Array): boolean => Buffer.from(a).equals(Buffer.from(b));

describe('real-world corpus', () => {
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
        // Unsupported/unsafe files must fail CLEANLY with a coded error,
        // never crash and never produce corrupt output.
        expect(error, `unexpected error: ${String(error)}`).toBeInstanceOf(SheetLiteError);
        return;
      }

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
