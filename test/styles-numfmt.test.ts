import { describe, it, expect } from 'vitest';
import * as XLSX from 'xlsx';
import { strFromU8 } from 'fflate';
import { Workbook } from '../src/index';
import { makeMinimalXlsx } from './fixtures/make';

/**
 * Regression: applying a CUSTOM number format (which must create the
 * `<numFmts>` container), a font and a fill in the same write used to collide
 * at the `<fonts>` offset and drop the `<numFmts>` element, leaving cellXfs
 * pointing at a numFmtId with no format code.
 */
describe('styles regression — <numFmts> container + fonts/fills in one write', () => {
  it('creates <numFmts> and keeps fonts/fills/counts correct', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const cell = wb.getWorksheet('Sheet1')!.getCell('B1');
    cell.numFmt = 'yyyy-mm-dd'; // custom (not built-in) -> container must be created
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: 'FFFF0000' };
    cell.value = new Date(Date.UTC(2026, 0, 15));

    const out = wb.toBuffer();
    const styles = strFromU8(Workbook.read(out).parts().get('xl/styles.xml')!);

    expect(styles).toContain('<numFmts');
    expect(styles).toContain('formatCode="yyyy-mm-dd"');
    expect(styles).toContain('<b/>');

    // Counts must match the real number of children (Excel validates them).
    expect(styles).toContain('<numFmts count="1"');
    expect(styles).toContain('<fonts count="2"');
    expect(styles).toContain('<fills count="2"');
    expect(styles).toContain('<cellXfs count="2"');

    const reread = Workbook.read(out).getWorksheet('Sheet1')!;
    expect(reread.getCell('B1').numFmt).toBe('yyyy-mm-dd');
    expect(reread.getCell('B1').font.bold).toBe(true);
    expect(reread.getCell('B1').fill).toMatchObject({ pattern: 'solid', fgColor: 'FFFF0000' });
    expect((reread.getCell('B1').value as Date).toISOString()).toBe('2026-01-15T00:00:00.000Z');
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('numFmt with quotes does not break the XML attribute', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    wb.getWorksheet('Sheet1')!.getCell('B2').numFmt = '#,##0.00 "€"';
    const out = wb.toBuffer();
    const styles = strFromU8(Workbook.read(out).parts().get('xl/styles.xml')!);
    // Quotes in the format code must be XML-escaped in the attribute...
    expect(styles).toContain('formatCode="#,##0.00 &quot;€&quot;"');
    // ...and decoded back when read.
    expect(Workbook.read(out).getWorksheet('Sheet1')!.getCell('B2').numFmt).toBe('#,##0.00 "€"');
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });
});
