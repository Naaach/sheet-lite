import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/index';
import { makeMinimalXlsx, EXPECTED_PART_NAMES } from './fixtures/make';

const bytesEqual = (a: Uint8Array, b: Uint8Array): boolean =>
  Buffer.from(a).equals(Buffer.from(b));

describe('v0.1 — read + lossless write (patch-only)', () => {
  it('reads every expected part of the package', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    expect(wb.parts().names().sort()).toEqual([...EXPECTED_PART_NAMES].sort());
  });

  it('no-op round-trip returns the original bytes exactly', () => {
    const original = makeMinimalXlsx();
    const out = Workbook.read(original).toBuffer();
    expect(bytesEqual(out, original)).toBe(true);
  });

  it('modifying one part keeps every other part byte-identical', () => {
    const original = makeMinimalXlsx();
    const wb = Workbook.read(original);
    const parts = wb.parts();

    const sheet = parts.get('xl/worksheets/sheet1.xml')!;
    const modified = new TextEncoder().encode(
      new TextDecoder().decode(sheet).replace('Hello', 'Hola'),
    );
    parts.set('xl/worksheets/sheet1.xml', modified);

    const out = wb.toBuffer();
    expect(bytesEqual(out, original)).toBe(false);

    const reread = Workbook.read(out).parts();
    expect(new TextDecoder().decode(reread.get('xl/worksheets/sheet1.xml')!)).toContain('Hola');

    for (const name of parts.names()) {
      if (name === 'xl/worksheets/sheet1.xml') continue;
      expect(bytesEqual(parts.get(name)!, reread.get(name)!), `part changed: ${name}`).toBe(true);
    }
  });

  it('reads content types from the package', () => {
    const ct = Workbook.read(makeMinimalXlsx()).contentTypes();
    expect(ct.defaults.get('xml')).toBe('application/xml');
    expect(ct.overrides.get('/xl/workbook.xml')).toContain('spreadsheetml.sheet.main');
  });
});
