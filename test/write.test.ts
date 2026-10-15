import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { mkdtempSync, readFileSync, rmSync, existsSync, readdirSync, writeFileSync, mkdirSync, statSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { inflateSync } from 'fflate';
import { Workbook, SheetLiteError, crc32, type CellValue } from '../src/index';
import { WorkbookWriter, type WriteRow } from '../src/write';
import { makeXlsxWithSheet } from './fixtures/make';

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

/** ExcelJS's load typings use a legacy global Buffer interface. */
type ExcelLoadInput = Parameters<ExcelJS.Workbook['xlsx']['load']>[0];

const tempDir = (): string => mkdtempSync(join(tmpdir(), 'sheet-lite-write-'));

/** Build a workbook in memory and return its bytes. */
async function build(
  fill: (wb: WorkbookWriter) => void | Promise<void>,
  options?: Parameters<typeof WorkbookWriter.create>[0],
): Promise<Uint8Array> {
  const wb = await WorkbookWriter.create(options);
  await fill(wb);
  return wb.toBuffer();
}

/** Cells of a sheet as `[ref, value]` pairs, through the eager reader. */
function readCells(bytes: Uint8Array, sheet: string): Array<[string, unknown]> {
  const ws = Workbook.read(bytes).getWorksheet(sheet)!;
  const out: Array<[string, unknown]> = [];
  ws.eachRow((row) => row.eachCell((cell) => out.push([cell.ref, cell.value])));
  return out;
}

function sheetXmlOf(bytes: Uint8Array, partName = 'xl/worksheets/sheet1.xml'): string {
  const store = Workbook.read(bytes).parts();
  return new TextDecoder().decode(store.get(partName)!);
}

interface CentralEntry {
  name: string;
  method: number;
  flags: number;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
  localOffset: number;
}

/**
 * Walk the ZIP central directory with an INDEPENDENT parser (not the code
 * under test), so the streaming writer's headers, CRC-32, sizes and offsets are
 * verified from the bytes alone.
 */
function readCentralDirectory(zip: Uint8Array): { entries: CentralEntry[]; eocdOffset: number; commentLen: number } {
  const dv = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  let eocd = zip.length - 22;
  while (eocd >= 0 && dv.getUint32(eocd, true) !== 0x0605_4b50) eocd -= 1;
  expect(eocd, 'EOCD record not found').toBeGreaterThanOrEqual(0);
  const total = dv.getUint16(eocd + 10, true);
  let pos = dv.getUint32(eocd + 16, true);
  const entries: CentralEntry[] = [];
  for (let i = 0; i < total; i += 1) {
    expect(dv.getUint32(pos, true)).toBe(0x0201_4b50); // central file header signature
    const nameLen = dv.getUint16(pos + 28, true);
    const extraLen = dv.getUint16(pos + 30, true);
    const commentLen = dv.getUint16(pos + 32, true);
    entries.push({
      name: new TextDecoder().decode(zip.subarray(pos + 46, pos + 46 + nameLen)),
      method: dv.getUint16(pos + 10, true),
      flags: dv.getUint16(pos + 8, true),
      crc: dv.getUint32(pos + 16, true),
      compressedSize: dv.getUint32(pos + 20, true),
      uncompressedSize: dv.getUint32(pos + 24, true),
      localOffset: dv.getUint32(pos + 42, true),
    });
    pos += 46 + nameLen + extraLen + commentLen;
  }
  return { entries, eocdOffset: eocd, commentLen: dv.getUint16(eocd + 20, true) };
}

/** Decompress an entry from its local header (independent of the writer). */
function readEntryBytes(zip: Uint8Array, entry: CentralEntry): Uint8Array {
  const dv = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  expect(dv.getUint32(entry.localOffset, true)).toBe(0x0403_4b50);
  const nameLen = dv.getUint16(entry.localOffset + 26, true);
  const extraLen = dv.getUint16(entry.localOffset + 28, true);
  const start = entry.localOffset + 30 + nameLen + extraLen;
  const payload = zip.subarray(start, start + entry.compressedSize);
  return entry.method === 0 ? payload.slice() : inflateSync(payload);
}

/** Excel/SheetJS-compatible value comparison (dates compare as ISO strings). */
function plain(value: unknown): unknown {
  if (value instanceof Date) return value.toISOString();
  if (value !== null && typeof value === 'object' && 'richText' in (value as object)) {
    return (value as { richText: Array<{ text: string }> })
      .richText.map((run) => run.text)
      .join('');
  }
  if (value !== null && typeof value === 'object' && 'formula' in (value as object)) {
    return { formula: (value as { formula: string }).formula };
  }
  return value;
}

// ---------------------------------------------------------------------------
// values
// ---------------------------------------------------------------------------

describe('write — round-trip through our own reader', () => {
  it('creates a minimal workbook that re-opens with its sheet and values', async () => {
    const bytes = await build((wb) => {
      const ws = wb.addWorksheet('Sheet1');
      ws.addRow(['Hello', 42]);
      ws.addRow(['World', 3.14]);
    });

    const wb = Workbook.read(bytes);
    expect(wb.sheetNames).toEqual(['Sheet1']);
    const ws = wb.getWorksheet('Sheet1')!;
    expect(ws.rowCount).toBe(2);
    expect(ws.getCell('A1').value).toBe('Hello');
    expect(ws.getCell('B1').value).toBe(42);
    expect(ws.getCell('B2').value).toBe(3.14);
  });

  it('round-trips every value kind', async () => {
    const richText = { richText: [{ text: 'Total: ' }, { text: '1 234 €', font: { bold: true } }] };
    const bytes = await build((wb) => {
      const ws = wb.addWorksheet('Kinds');
      const row = ws.addRow([
        'text',
        -12.5,
        true,
        false,
        null,
        new Date(Date.UTC(2024, 4, 17, 6, 30, 15)),
        richText,
        { formula: 'SUM(B1:B1)', result: -12.5 },
        { error: '#REF!' },
        '',
      ]);
      row.getCell(6).numFmt = 'yyyy-mm-dd hh:mm';
    });

    const ws = Workbook.read(bytes).getWorksheet('Kinds')!;
    expect(ws.getCell('A1').value).toBe('text');
    expect(ws.getCell('B1').value).toBe(-12.5);
    expect(ws.getCell('C1').value).toBe(true);
    expect(ws.getCell('D1').value).toBe(false);
    expect(ws.getCell('E1').value).toBeNull();
    expect(ws.getCell('F1').value).toBeInstanceOf(Date);
    expect((ws.getCell('F1').value as Date).toISOString()).toBe('2024-05-17T06:30:15.000Z');
    expect(ws.getCell('G1').value).toEqual(richText);
    expect(ws.getCell('H1').value).toEqual({ formula: 'SUM(B1:B1)', result: -12.5 });
    expect(ws.getCell('I1').value).toEqual({ error: '#REF!' });
    expect(ws.getCell('J1').value).toBe('');
  });

  it('escapes XML specials and preserves surrounding whitespace', async () => {
    const nasty = 'a & b < c > d " e \' f';
    const bytes = await build((wb) => {
      const ws = wb.addWorksheet('Escapes');
      ws.addRow([nasty, '  padded  ', 'line\nbreak\ttab']);
    });
    const xml = sheetXmlOf(bytes);
    expect(xml).toContain('&amp;');
    expect(xml).toContain('&lt;');
    expect(xml).toContain('&quot;');
    expect(xml).toContain('&apos;');
    // Every string with leading/trailing whitespace or a newline carries
    // xml:space="preserve".
    expect(xml).toBe(xml.replace(/<t xml:space="preserve">/g, '<t xml:space="preserve">'));
    expect((xml.match(/xml:space="preserve"/g) ?? []).length).toBe(2);

    const ws = Workbook.read(bytes).getWorksheet('Escapes')!;
    expect(ws.getCell('A1').value).toBe(nasty);
    expect(ws.getCell('B1').value).toBe('  padded  ');
    expect(ws.getCell('C1').value).toBe('line\nbreak\ttab');
  });

  it('writes empty rows, sparse rows and columns beyond Z', async () => {
    const bytes = await build((wb) => {
      const ws = wb.addWorksheet('Wide');
      ws.addRow([]);
      ws.addRow(['a']);
      const row = ws.addRow([]);
      row.getCell(27).value = 'AA'; // 27 -> AA
      row.getCell(30).value = 'AD'; // 30 -> AD
      row.getCell(31).value = 'AE';
    });
    const ws = Workbook.read(bytes).getWorksheet('Wide')!;
    expect(ws.rowCount).toBe(3);
    expect(ws.columnCount).toBe(31);
    expect(ws.getCell('AA3').value).toBe('AA');
    expect(ws.getCell('AD3').value).toBe('AD');
    expect(ws.getCell('AE3').value).toBe('AE');
    expect(ws.getCell('A2').value).toBe('a');
    expect(ws.getCell('B2').value).toBeNull();
    // Row 1 and row 3 are empty rows: `<row r="1"/>`.
    expect(sheetXmlOf(bytes)).toContain('<row r="1"/>');
    expect(sheetXmlOf(bytes)).toContain('<row r="3">');

    // XFD (16384) is the last column Excel accepts.
    const last = await build((wb) => {
      wb.addWorksheet('Edge').addRow([]).getCell(16_384).value = 'XFD';
    });
    expect(Workbook.read(last).getWorksheet('Edge')!.getCell('XFD1').value).toBe('XFD');
  });

  it('writes the 1904 date system when asked', async () => {
    const date = new Date(Date.UTC(2024, 0, 1));
    const serial1900 = await build((wb) => {
      const row = wb.addWorksheet('D').addRow([date]);
      row.numFmt = 'yyyy-mm-dd';
    });
    const serial1904 = await build(
      (wb) => {
        const row = wb.addWorksheet('D').addRow([date]);
        row.numFmt = 'yyyy-mm-dd';
      },
      { date1904: true },
    );

    const a = Workbook.read(serial1900).getWorksheet('D')!.getCell('A1').value as Date;
    const b = Workbook.read(serial1904).getWorksheet('D')!.getCell('A1').value as Date;
    expect(a.toISOString()).toBe(date.toISOString());
    expect(b.toISOString()).toBe(date.toISOString());
    // The same instant is 1462 serials lower on the 1904 system.
    const s1900 = Number(sheetXmlOf(serial1900).match(/<v>([\d.]+)<\/v>/)![1]);
    const s1904 = Number(sheetXmlOf(serial1904).match(/<v>([\d.]+)<\/v>/)![1]);
    expect(s1900 - s1904).toBe(1462);
  });

  it('keeps full double precision', async () => {
    const bytes = await build((wb) => {
      wb.addWorksheet('N').addRow([0.1 + 0.2, 1 / 3, 1e-7, 123456789.123456789]);
    });
    const ws = Workbook.read(bytes).getWorksheet('N')!;
    expect(ws.getCell('A1').value).toBe(0.1 + 0.2);
    expect(ws.getCell('B1').value).toBe(1 / 3);
    expect(ws.getCell('C1').value).toBe(1e-7);
    expect(ws.getCell('D1').value).toBe(123456789.123456789);
  });

  it('writes several sheets with coherent registration', async () => {
    const bytes = await build((wb) => {
      const first = wb.addWorksheet('First');
      first.addRow([1]);
      first.addRow([2]);
      const second = wb.addWorksheet('Second');
      second.addRow(['x']);
      // A sheet created but never written to still has to exist in the package.
      wb.addWorksheet('Empty');
    });

    const wb = Workbook.read(bytes);
    expect(wb.sheetNames).toEqual(['First', 'Second', 'Empty']);
    expect(wb.getWorksheet(1)!.getCell('A2').value).toBe(2);
    expect(wb.getWorksheet('Second')!.getCell('A1').value).toBe('x');
    expect(wb.getWorksheet('Empty')!.rowCount).toBe(0);

    const store = wb.parts();
    expect(store.has('xl/worksheets/sheet1.xml')).toBe(true);
    expect(store.has('xl/worksheets/sheet2.xml')).toBe(true);
    expect(store.has('xl/worksheets/sheet3.xml')).toBe(true);

    // Every sheet is registered in workbook.xml, in a relationship and in the
    // content types — no dangling reference in either direction.
    const workbookXml = new TextDecoder().decode(store.get('xl/workbook.xml')!);
    const rels = new TextDecoder().decode(store.get('xl/_rels/workbook.xml.rels')!);
    const ct = new TextDecoder().decode(store.get('[Content_Types].xml')!);
    for (const [index, name] of ['First', 'Second', 'Empty'].entries()) {
      expect(workbookXml).toContain(`name="${name}"`);
      expect(workbookXml).toContain(`r:id="rId${index + 1}"`);
      expect(rels).toContain(`Id="rId${index + 1}"`);
      expect(rels).toContain(`worksheets/sheet${index + 1}.xml`);
      expect(ct).toContain(`/xl/worksheets/sheet${index + 1}.xml`);
    }
    expect(ct).toContain('/xl/workbook.xml');
    expect(ct).toContain('/xl/styles.xml');
    expect(new TextDecoder().decode(store.get('_rels/.rels')!)).toContain('xl/workbook.xml');
  });
});

// ---------------------------------------------------------------------------
// styles and sheet structure
// ---------------------------------------------------------------------------

describe('write — styles and sheet structure', () => {
  it('applies numFmt, font, fill, border and alignment', async () => {
    const bytes = await build((wb) => {
      const ws = wb.addWorksheet('Styled');
      const row = ws.addRow(['Total', 1234.5]);
      row.getCell(2).numFmt = '#,##0.00';
      row.getCell(2).font = { bold: true, size: 14, color: 'FFFF0000', name: 'Arial' };
      row.getCell(2).numFmt = '#,##0.00 "EUR"'; // not a built-in format -> custom id
      row.getCell(1).fill = { pattern: 'solid', fgColor: 'FFFFFF00' };
      row.getCell(1).border = {
        top: { style: 'thin', color: { argb: 'FF00B0F0' } },
        bottom: { style: 'double' },
      };
      row.getCell(1).alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
      row.getCell(1).font = { italic: true }; // merges over the previous font patch
    });

    const ws = Workbook.read(bytes).getWorksheet('Styled')!;
    expect(ws.getCell('B1').numFmt).toBe('#,##0.00 "EUR"');
    expect(ws.getCell('B1').font).toEqual({ bold: true, size: 14, color: 'FFFF0000', name: 'Arial' });
    expect(ws.getCell('A1').fill).toMatchObject({ pattern: 'solid', fgColor: 'FFFFFF00' });
    expect(ws.getCell('A1').border).toMatchObject({
      top: { style: 'thin', color: { argb: 'FF00B0F0' } },
      bottom: { style: 'double' },
    });
    expect(ws.getCell('A1').alignment).toMatchObject({ horizontal: 'center', wrapText: true });
    expect(ws.getCell('A1').font).toMatchObject({ italic: true });

    // Identical styles collapse to one entry: a 1-row sheet needs 1 default xf
    // plus the two it actually uses.
    const styles = new TextDecoder().decode(Workbook.read(bytes).parts().get('xl/styles.xml')!);
    const count = Number(/<cellXfs count="(\d+)"/.exec(styles)![1]);
    expect(count).toBeLessThanOrEqual(4);
    expect(styles).toContain('<numFmt numFmtId="164" formatCode="#,##0.00 &quot;EUR&quot;"/>');
    // Every count attribute agrees with the number of children it holds.
    expect(Number(/<fonts count="(\d+)"/.exec(styles)![1])).toBe(
      (styles.match(/<font>/g) ?? []).length,
    );
    expect(Number(/<borders count="(\d+)"/.exec(styles)![1])).toBe(
      (styles.match(/<border[/>]/g) ?? []).length,
    );
    expect(Number(/<fills count="(\d+)"/.exec(styles)![1])).toBe(
      (styles.match(/<fill>/g) ?? []).length,
    );
  });

  it('writes an empty cell that only carries a style', async () => {
    const bytes = await build((wb) => {
      const row = wb.addWorksheet('Styled').addRow([null]);
      row.getCell(1).fill = { pattern: 'solid', fgColor: 'FF00FF00' };
    });
    const xml = sheetXmlOf(bytes);
    expect(xml).toMatch(/<c r="A1" s="\d+"\/>/);
    expect(Workbook.read(bytes).getWorksheet('Styled')!.getCell('A1').value).toBeNull();
  });

  it('writes column widths, hidden columns, row heights, frozen panes and merges', async () => {
    const bytes = await build((wb) => {
      const ws = wb.addWorksheet('Layout', { views: [{ state: 'frozen', xSplit: 2, ySplit: 3 }] });
      ws.getColumn('A').width = 24;
      ws.getColumn('C').hidden = true;
      ws.mergeCells('A1:C1');
      const row = ws.addRow(['a', 'b', 'c']);
      row.height = 22;
    });

    const ws = Workbook.read(bytes).getWorksheet('Layout')!;
    expect(ws.getColumn('A').width).toBe(24);
    expect(ws.getColumn('B').width).toBeUndefined();
    expect(ws.getColumn('C').hidden).toBe(true);
    expect(ws.getRow(1).height).toBe(22);
    expect(ws.views[0]).toMatchObject({ state: 'frozen', xSplit: 2, ySplit: 3, topLeftCell: 'C4' });
    expect(sheetXmlOf(bytes)).toContain('<mergeCell ref="A1:C1"/>');
  });

  it('applies a row-wide style to every cell of the row', async () => {
    const bytes = await build((wb) => {
      const header = wb.addWorksheet('Head').addRow(['name', 'amount']);
      header.font = { bold: true, color: 'FFFFFFFF' };
      header.fill = { pattern: 'solid', fgColor: 'FF203864' };
      header.numFmt = '@';
      header.height = 30;
    });
    const ws = Workbook.read(bytes).getWorksheet('Head')!;
    for (const ref of ['A1', 'B1']) {
      expect(ws.getCell(ref).font).toMatchObject({ bold: true, color: 'FFFFFFFF' });
      expect(ws.getCell(ref).fill).toMatchObject({ fgColor: 'FF203864' });
      expect(ws.getCell(ref).numFmt).toBe('@');
    }
    // One shared xf for the whole row.
    const styles = new TextDecoder().decode(Workbook.read(bytes).parts().get('xl/styles.xml')!);
    expect(Number(/<cellXfs count="(\d+)"/.exec(styles)![1])).toBe(2);
  });

  it('freezes panes through the `views` setter and reports them back', async () => {
    const bytes = await build((wb) => {
      const ws = wb.addWorksheet('Frozen');
      ws.views = [{ state: 'frozen', ySplit: 1 }];
      ws.addRow(['header']);
    });
    expect(Workbook.read(bytes).getWorksheet('Frozen')!.views).toMatchObject([
      { state: 'frozen', ySplit: 1, topLeftCell: 'A2' },
    ]);
  });

  it('keeps reporting the accepted views after a rejected set (PROLOGUE_LOCKED)', async () => {
    // The setter used to assign `viewsValue` BEFORE `stream.setViews()` could
    // throw, so a locked sheet reported panes that were never written.
    const wb = await WorkbookWriter.create();
    const ws = wb.addWorksheet('Locked', { views: [{ state: 'frozen', ySplit: 1 }] });
    expect(ws.views).toMatchObject([{ state: 'frozen', ySplit: 1 }]);

    ws.addRow(['a']);
    ws.addRow(['b']); // flushes row 1: the prologue is emitted, views are locked
    try {
      ws.views = [{ state: 'frozen', xSplit: 4 }];
      expect.unreachable();
    } catch (error) {
      expect((error as SheetLiteError).code).toBe('PROLOGUE_LOCKED');
    }
    // The rejected value was never written anywhere, so the getter still
    // reports the one that is in the part.
    expect(ws.views).toMatchObject([{ state: 'frozen', ySplit: 1 }]);

    const bytes = await wb.toBuffer();
    const xml = sheetXmlOf(bytes);
    expect(xml).toContain('<pane ySplit="1"');
    expect(xml).not.toContain('xSplit="4"');
    expect(Workbook.read(bytes).getWorksheet('Locked')!.views).toMatchObject([
      { state: 'frozen', ySplit: 1 },
    ]);
  });

  it('rejects changes to a row that was already written (ROW_WRITTEN)', async () => {
    const wb = await WorkbookWriter.create();
    const ws = wb.addWorksheet('S');
    const first = ws.addRow(['a']);
    ws.addRow(['b']); // flushes the first row
    expect(() => {
      first.getCell(1).value = 'late';
    }).toThrow(SheetLiteError);
    for (const change of [
      () => {
        first.height = 10;
      },
      () => {
        first.getCell(1).font = { bold: true };
      },
      () => {
        first.numFmt = '0.00';
      },
    ]) {
      let code = '';
      try {
        change();
      } catch (error) {
        code = (error as SheetLiteError).code;
      }
      expect(code).toBe('ROW_WRITTEN');
    }
    await wb.toBuffer();
  });

  it('locks the sheet prologue once the part has started (PROLOGUE_LOCKED)', async () => {
    const wb = await WorkbookWriter.create();
    const ws = wb.addWorksheet('S');
    // The first row is still buffered, so the prologue is still open.
    ws.addRow(['a']);
    expect(() => {
      ws.getColumn('B').width = 10;
    }).not.toThrow();
    ws.addRow(['b']); // flushes row 1: the prologue is now emitted
    let code = '';
    try {
      ws.getColumn('C').width = 10;
    } catch (error) {
      code = (error as SheetLiteError).code;
    }
    expect(code).toBe('PROLOGUE_LOCKED');
    try {
      ws.views = [{ state: 'frozen', ySplit: 1 }];
      expect.unreachable();
    } catch (error) {
      expect((error as SheetLiteError).code).toBe('PROLOGUE_LOCKED');
    }
    await wb.toBuffer();
  });
});

// ---------------------------------------------------------------------------
// ZIP container
// ---------------------------------------------------------------------------

describe('write — parity with the eager (patch) writer', () => {
  /**
   * The incremental writer ships its own copy of the cell serializer so the
   * gzip-budgeted main entry never statically imports the writer's kernel
   * (see `writer/cell-markup.ts`). These tests pin the two implementations
   * together: identical values must produce identical markup and identical
   * style-sheet entries, whichever writer emitted them.
   */
  const VALUES: CellValue[] = [
    'plain',
    '  padded  ',
    'a & b < c > "d" \'e\'',
    'line\nbreak\ttab',
    0,
    -12.5,
    1 / 3,
    1e-7,
    true,
    false,
    new Date(Date.UTC(2024, 4, 17, 6, 30, 15)),
    { error: '#DIV/0!' },
    { formula: 'SUM(A1:A2)', result: 7 },
    { formula: 'CONCAT("a","b")', result: 'ab' },
    { formula: 'B1>0', result: true },
    { richText: [{ text: 'Total: ' }, { text: '1 234 €', font: { bold: true, color: 'FFFF0000' } }] },
    { richText: [] },
  ];

  it('emits the same <c> markup for every value kind', async () => {
    const streamed = await build((wb) => {
      wb.addWorksheet('Parity').addRow(VALUES);
    });
    const eager = (() => {
      const wb = Workbook.read(makeXlsxWithSheet(
        '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          '<sheetData><row r="1"/></sheetData></worksheet>',
      ));
      const ws = wb.getWorksheet('Sheet1')!;
      VALUES.forEach((value, index) => {
        ws.getCell(1, index + 1).value = value;
      });
      return wb.toBuffer();
    })();

    const cellsOf = (bytes: Uint8Array): string[] => {
      const xml = sheetXmlOf(bytes, 'xl/worksheets/sheet1.xml');
      return [...xml.matchAll(/<c [^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)].map((m) => m[0]);
    };
    const streamedCells = cellsOf(streamed);
    const eagerCells = cellsOf(eager);
    // `{ richText: [] }` renders as `<is><t/></is>` in both writers; nothing is
    // dropped, so both must produce one <c> per value.
    expect(streamedCells).toHaveLength(VALUES.length);
    expect(eagerCells).toHaveLength(VALUES.length);
    expect(streamedCells).toEqual(eagerCells);
  });

  it('registers the same style entries in styles.xml', async () => {
    const patch = (row: WriteRow): void => {
      row.font = { bold: true };
      row.getCell(2).numFmt = '#,##0.00';
      row.getCell(3).fill = { pattern: 'solid', fgColor: 'FFFFFF00' };
      row.getCell(4).border = { top: { style: 'thin', color: { argb: 'FF00B0F0' } } };
      row.getCell(5).alignment = { horizontal: 'center', wrapText: true };
    };

    const streamed = await build((wb) => {
      const ws = wb.addWorksheet('Styled');
      patch(ws.addRow(['a', 1, 'b', 'c', 'd']));
      patch(ws.addRow(['a', 1, 'b', 'c', 'd'])); // identical styles dedup
    });
    const eager = (() => {
      const wb = Workbook.read(makeXlsxWithSheet(
        '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          '<sheetData><row r="1"/></sheetData></worksheet>',
      ));
      const ws = wb.getWorksheet('Sheet1')!;
      for (let r = 1; r <= 2; r += 1) {
        ws.getCell(r, 1).value = 'a';
        ws.getCell(r, 1).font = { bold: true };
        ws.getCell(r, 2).value = 1;
        ws.getCell(r, 2).font = { bold: true };
        ws.getCell(r, 2).numFmt = '#,##0.00';
        ws.getCell(r, 3).value = 'b';
        ws.getCell(r, 3).font = { bold: true };
        ws.getCell(r, 3).fill = { pattern: 'solid', fgColor: 'FFFFFF00' };
        ws.getCell(r, 4).value = 'c';
        ws.getCell(r, 4).font = { bold: true };
        ws.getCell(r, 4).border = { top: { style: 'thin', color: { argb: 'FF00B0F0' } } };
        ws.getCell(r, 5).value = 'd';
        ws.getCell(r, 5).font = { bold: true };
        ws.getCell(r, 5).alignment = { horizontal: 'center', wrapText: true };
      }
      return wb.toBuffer();
    })();

    // Identical containers and entries: the writer's baseline matches the one
    // the eager path creates for a file without styles.xml, so ids match too.
    const containers = (bytes: Uint8Array): string[] => {
      const xml = new TextDecoder().decode(Workbook.read(bytes).parts().get('xl/styles.xml')!);
      return [...xml.matchAll(/<(numFmt|font|fill|border|xf)\b[^>]*?(?:\/>|>[\s\S]*?<\/\1>)/g)].map(
        (m) => m[0]!.replace(/\s+xmlns(:\w+)?="[^"]*"/g, ''),
      );
    };
    expect(containers(streamed)).toEqual(containers(eager));
  });
});

describe('write — ZIP container', () => {
  it('emits correct local headers, data descriptors, CRCs and the central directory', async () => {
    const bytes = await build((wb) => {
      wb.addWorksheet('Data').addRow(['hello', 1]);
    });
    const { entries, eocdOffset, commentLen } = readCentralDirectory(bytes);
    expect(entries.map((entry) => entry.name).sort()).toEqual([
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/_rels/workbook.xml.rels',
      'xl/styles.xml',
      'xl/workbook.xml',
      'xl/worksheets/sheet1.xml',
    ]);
    expect(commentLen).toBe(0);
    expect(bytes.length).toBe(eocdOffset + 22);

    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    for (const entry of entries) {
      expect(entry.localOffset).toBeLessThan(eocdOffset);
      // The streaming writer cannot know the sizes up front: bit 3 is set and
      // the local header carries zeroes, the real values follow the data.
      expect(entry.flags & 0x0008, `${entry.name} data-descriptor flag`).toBe(0x0008);
      expect(dv.getUint16(entry.localOffset + 6, true) & 0x0008).toBe(0x0008);
      expect(dv.getUint32(entry.localOffset + 14, true)).toBe(0); // crc in local header
      expect(dv.getUint32(entry.localOffset + 18, true)).toBe(0); // compressed size
      expect(dv.getUint32(entry.localOffset + 22, true)).toBe(0); // uncompressed size

      const data = readEntryBytes(bytes, entry);
      expect(data.length).toBe(entry.uncompressedSize);
      expect(crc32(data)).toBe(entry.crc);

      // The data descriptor itself: signature + CRC + both sizes.
      const descriptorAt = entry.localOffset + 30 +
        dv.getUint16(entry.localOffset + 26, true) +
        dv.getUint16(entry.localOffset + 28, true) +
        entry.compressedSize;
      expect(dv.getUint32(descriptorAt, true)).toBe(0x0807_4b50);
      expect(dv.getUint32(descriptorAt + 4, true)).toBe(entry.crc);
      expect(dv.getUint32(descriptorAt + 8, true)).toBe(entry.compressedSize);
      expect(dv.getUint32(descriptorAt + 12, true)).toBe(entry.uncompressedSize);
      expect(descriptorAt + 16).toBeLessThanOrEqual(eocdOffset);
    }

    // EOCD agrees with the directory: the central directory starts where the
    // EOCD says and is exactly as long as it claims.
    expect(dv.getUint16(eocdOffset + 8, true)).toBe(entries.length);
    expect(dv.getUint16(eocdOffset + 10, true)).toBe(entries.length);
    const centralSize = dv.getUint32(eocdOffset + 12, true);
    const centralOffset = dv.getUint32(eocdOffset + 16, true);
    expect(centralOffset + centralSize).toBe(eocdOffset);
    expect(dv.getUint32(centralOffset, true)).toBe(0x0201_4b50);
  });

  it('writes stored entries at level 0', async () => {
    const bytes = await build((wb) => {
      wb.addWorksheet('Plain').addRow(['x', 1]);
    }, { level: 0 });
    const { entries } = readCentralDirectory(bytes);
    for (const entry of entries) {
      expect(entry.method).toBe(0);
      expect(entry.compressedSize).toBe(entry.uncompressedSize);
    }
    expect(Workbook.read(bytes).getWorksheet('Plain')!.getCell('A1').value).toBe('x');
  });

  it('re-reads a streamed file through the streaming reader (bounded path)', async () => {
    const dir = tempDir();
    const path = join(dir, 'streamed.xlsx');
    try {
      const wb = await Workbook.create({ file: path });
      const ws = wb.addWorksheet('Data');
      for (let i = 1; i <= 500; i += 1) ws.addRow([i, `row-${i}`]);
      await wb.commit();
      expect(existsSync(path)).toBe(true);

      const { StreamingWorkbook } = await import('../src/stream');
      const reader = await StreamingWorkbook.open(path);
      try {
        let rows = 0;
        for await (const row of reader.readRows('Data')) rows += 1;
        expect(rows).toBe(500);
      } finally {
        await reader.close();
      }
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

// ---------------------------------------------------------------------------
// finalization order: a ZIP entry's bytes must be contiguous
// ---------------------------------------------------------------------------

describe('write — sheet part finalization order', () => {
  /** Rows a sheet is expected to read back as, flattened cell by cell. */
  type ExpectedSheet = { name: string; rows: Array<Array<string | number>> };

  /**
   * A complete package, verified independently of the writer: every declared
   * sheet part exists in the archive, every entry's CRC matches its own bytes,
   * and every sheet reads back with exactly the rows it was given.
   */
  function expectCompletePackage(bytes: Uint8Array, expected: ExpectedSheet[]): void {
    const { entries } = readCentralDirectory(bytes);
    for (const entry of entries) {
      expect(crc32(readEntryBytes(bytes, entry)), `${entry.name} CRC`).toBe(entry.crc);
    }
    const wb = Workbook.read(bytes);
    expect(wb.sheetNames).toEqual(expected.map((sheet) => sheet.name));
    for (const [index, sheet] of expected.entries()) {
      expect(wb.parts().has(`xl/worksheets/sheet${index + 1}.xml`), sheet.name).toBe(true);
      expect(wb.getWorksheet(sheet.name)!.rowCount, `${sheet.name} rowCount`).toBe(sheet.rows.length);
      expect(
        readCells(bytes, sheet.name).map(([, value]) => plain(value)),
        `${sheet.name} cells`,
      ).toEqual(sheet.rows.flat());
    }
  }

  it('commits when an early sheet holds an open entry and a later one is empty', async () => {
    // The regression: sheet A flushed a row, so `begin()` already emitted its
    // ZIP local header and its entry is OPEN. Sheet B is created but never
    // written to, so it has no entry yet — closing it first would make
    // `begin()` silently close A's entry and strand A's still-buffered row
    // (BAD_STATE: "No open ZIP entry").
    const bytes = await build((wb) => {
      const a = wb.addWorksheet('A');
      a.addRow([1, 'x']);
      a.addRow([2, 'y']);
      wb.addWorksheet('B');
    });
    expectCompletePackage(bytes, [
      { name: 'A', rows: [[1, 'x'], [2, 'y']] },
      { name: 'B', rows: [] },
    ]);
  });

  it('commits the mirror case: only the LAST sheet has flushed its entry', async () => {
    const bytes = await build((wb) => {
      wb.addWorksheet('First'); // created, never written to
      const last = wb.addWorksheet('Last');
      last.addRow(['a']);
      last.addRow(['b']);
    });
    expectCompletePackage(bytes, [
      { name: 'First', rows: [] },
      { name: 'Last', rows: [['a'], ['b']] },
    ]);
  });

  it('commits three sheets with mixed row counts', async () => {
    const bytes = await build((wb) => {
      const a = wb.addWorksheet('A');
      a.addRow([1]);
      a.addRow([2]);
      a.addRow([3]);
      wb.addWorksheet('B'); // empty, never activated
      const c = wb.addWorksheet('C');
      c.addRow(['c1']);
      c.addRow(['c2']);
    });
    expectCompletePackage(bytes, [
      { name: 'A', rows: [[1], [2], [3]] },
      { name: 'B', rows: [] },
      { name: 'C', rows: [['c1'], ['c2']] },
    ]);
  });

  it('commits when only the MIDDLE sheet holds the open entry', async () => {
    // The active sheet is neither first nor last: both empty sheets still have
    // to be emitted, and neither may steal the open entry.
    const bytes = await build((wb) => {
      wb.addWorksheet('A');
      const b = wb.addWorksheet('B');
      b.addRow(['b1']);
      b.addRow(['b2']);
      wb.addWorksheet('C');
    });
    expectCompletePackage(bytes, [
      { name: 'A', rows: [] },
      { name: 'B', rows: [['b1'], ['b2']] },
      { name: 'C', rows: [] },
    ]);
  });

  it('is idempotent: a second commit() keeps the archive intact', async () => {
    const wb = await WorkbookWriter.create();
    const a = wb.addWorksheet('A');
    a.addRow([1]);
    a.addRow([2]);
    wb.addWorksheet('B');
    await wb.commit();
    await wb.commit();
    const bytes = await wb.toBuffer();
    expectCompletePackage(bytes, [
      { name: 'A', rows: [[1], [2]] },
      { name: 'B', rows: [] },
    ]);
  });

  it('emits the prologue of the sheet that flushed first before its first row', async () => {
    // Closing the right entry in the right order must not move the prologue:
    // `<sheetViews>`/`<cols>` precede `<sheetData>` even for the sheet whose
    // entry was opened first and is closed first.
    const bytes = await build((wb) => {
      const a = wb.addWorksheet('A', { views: [{ state: 'frozen', ySplit: 1 }] });
      a.getColumn('B').width = 20;
      a.addRow([1]);
      a.addRow([2]);
      wb.addWorksheet('B', { views: [{ state: 'frozen', ySplit: 2 }] });
    });

    const xml = sheetXmlOf(bytes, 'xl/worksheets/sheet1.xml');
    expect(xml.indexOf('<sheetViews>')).toBeGreaterThan(-1);
    expect(xml.indexOf('<sheetViews>')).toBeLessThan(xml.indexOf('<sheetData>'));
    expect(xml.indexOf('<cols>')).toBeLessThan(xml.indexOf('<sheetData>'));
    expect(xml.indexOf('<row r="1"')).toBeGreaterThan(xml.indexOf('<sheetData>'));

    // The empty sheet is emitted with its own (later) prologue and no rows.
    const empty = sheetXmlOf(bytes, 'xl/worksheets/sheet2.xml');
    expect(empty).toContain('<sheetViews>');
    expect(empty).toContain('<sheetData></sheetData>');
    expect(Workbook.read(bytes).getWorksheet('B')!.views).toMatchObject([
      { state: 'frozen', ySplit: 2 },
    ]);
  });

  it('writes the same complete package to a file target', async () => {
    const dir = tempDir();
    const path = join(dir, 'two-sheets.xlsx');
    try {
      const wb = await Workbook.create({ file: path });
      const a = wb.addWorksheet('A');
      a.addRow([1, 'x']);
      a.addRow([2, 'y']);
      wb.addWorksheet('B');
      await wb.commit();
      expectCompletePackage(new Uint8Array(readFileSync(path)), [
        { name: 'A', rows: [[1, 'x'], [2, 'y']] },
        { name: 'B', rows: [] },
      ]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

// ---------------------------------------------------------------------------
// lifecycle
// ---------------------------------------------------------------------------

describe('write — lifecycle and validation', () => {
  it('rejects invalid and duplicate sheet names', async () => {
    const wb = await WorkbookWriter.create();
    wb.addWorksheet('Ok');
    for (const name of ['', 'x'.repeat(32), 'a:b', 'a\\b', 'a/b', 'a?b', 'a*b', 'a[b', 'a]b']) {
      let code = '';
      try {
        wb.addWorksheet(name);
      } catch (error) {
        code = (error as SheetLiteError).code;
      }
      expect(code, `name ${JSON.stringify(name)}`).toBe('BAD_SHEET_NAME');
    }
    let code = '';
    try {
      wb.addWorksheet('Ok');
    } catch (error) {
      code = (error as SheetLiteError).code;
    }
    expect(code).toBe('SHEET_EXISTS');
    await wb.toBuffer();
  });

  it('refuses to write a workbook without sheets', async () => {
    const wb = await WorkbookWriter.create();
    await expect(wb.commit()).rejects.toThrow(/without worksheets/);
  });

  it('rejects writes after commit (SHEET_FINALIZED)', async () => {
    const wb = await WorkbookWriter.create();
    const ws = wb.addWorksheet('Done');
    ws.addRow(['a']);
    await wb.commit();
    expect(wb.sealed).toBe(true);
    expect(() => ws.addRow(['late'])).toThrow(/already been written/);
    expect(() => wb.addWorksheet('Other')).toThrow(/already been committed/);
    // commit() is idempotent and toBuffer() stays usable afterwards.
    await wb.commit();
    expect((await wb.toBuffer()).length).toBeGreaterThan(0);
  });

  it('rejects interleaved writes between two sheets', async () => {
    const wb = await WorkbookWriter.create();
    const a = wb.addWorksheet('A');
    const b = wb.addWorksheet('B');
    a.addRow(['a1']);
    b.addRow(['b1']);
    let code = '';
    try {
      a.addRow(['a2']);
    } catch (error) {
      code = (error as SheetLiteError).code;
    }
    expect(code).toBe('SHEET_FINALIZED'); // the sheet part is already closed
    const bytes = await wb.toBuffer();
    expect(Workbook.read(bytes).getWorksheet('A')!.getCell('A1').value).toBe('a1');
    expect(Workbook.read(bytes).getWorksheet('B')!.getCell('A1').value).toBe('b1');
  });

  it('rejects non-finite numbers, invalid dates and XML-illegal characters', async () => {
    const wb = await WorkbookWriter.create();
    const ws = wb.addWorksheet('Bad');
    expect(() => ws.addRow([Number.NaN])).toThrow(/BAD_NUMBER|not a valid cell number/);
    expect(() => ws.addRow([Number.POSITIVE_INFINITY])).toThrow(/not a valid cell number/);
    expect(() => ws.addRow([new Date('nope')])).toThrow(/Date is invalid/);
    expect(() => ws.addRow(['bad\u0007char'])).toThrow(/not valid in XML/);
    try {
      ws.addRow([{ richText: [{ text: 'x\u0000y' }] }]);
      expect.unreachable();
    } catch (error) {
      expect((error as SheetLiteError).code).toBe('BAD_TEXT');
    }
    expect(() => ws.addRow([{ formula: 'A1', result: Number.NaN }])).toThrow(/finite/);
    await wb.toBuffer();
  });

  it('rejects bad column and row heights', async () => {
    const wb = await WorkbookWriter.create();
    const ws = wb.addWorksheet('Limits');
    const row = ws.addRow(['a']);
    expect(() => row.height = -1).toThrow(/non-negative/);
    expect(() => ws.getColumn(0)).toThrow(/Invalid column/);
    expect(() => ws.getColumn(16_385)).toThrow(/Invalid column/);
    expect(() => ws.getColumn('!')).toThrow(/Invalid column/);
    expect(() => ws.addRow(new Array(16_385).fill('x'))).toThrow(/at most 16384 columns/);
    await wb.toBuffer();
  });

  it('streams to a file target and to a custom sink', async () => {
    const dir = tempDir();
    const filePath = join(dir, 'out.xlsx');
    try {
      const fileWriter = await WorkbookWriter.create({ file: filePath });
      fileWriter.addWorksheet('F').addRow(['from file']);
      await fileWriter.commit();
      const fileBytes = new Uint8Array(readFileSync(filePath));

      const chunks: Uint8Array[] = [];
      const sinkWriter = await WorkbookWriter.create({
        sink: { write: (chunk) => chunks.push(chunk) },
      });
      sinkWriter.addWorksheet('S').addRow(['from sink']);
      await sinkWriter.commit();
      const sinkBytes = readCentralDirectory(
        (() => {
          const total = chunks.reduce((n, chunk) => n + chunk.length, 0);
          const out = new Uint8Array(total);
          let at = 0;
          for (const chunk of chunks) {
            out.set(chunk, at);
            at += chunk.length;
          }
          return out;
        })(),
      );

      expect(fileBytes.length).toBeGreaterThan(0);
      expect(Workbook.read(fileBytes).getWorksheet('F')!.getCell('A1').value).toBe('from file');
      expect(sinkBytes.entries.map((entry) => entry.name)).toContain('xl/worksheets/sheet1.xml');
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('refuses toBuffer() on a streaming target (BAD_TARGET)', async () => {
    const dir = tempDir();
    try {
      const wb = await WorkbookWriter.create({ file: join(dir, 'streamed.xlsx') });
      wb.addWorksheet('S').addRow(['x']);
      let code = '';
      try {
        await wb.toBuffer();
      } catch (error) {
        code = (error as SheetLiteError).code;
      }
      expect(code).toBe('BAD_TARGET');
      await wb.commit();
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('writeFile() writes a buffered workbook to disk', async () => {
    const dir = tempDir();
    const path = join(dir, 'written.xlsx');
    try {
      const wb = await WorkbookWriter.create();
      wb.addWorksheet('W').addRow(['written']);
      await wb.writeFile(path);
      const bytes = new Uint8Array(readFileSync(path));
      expect(Workbook.read(bytes).getWorksheet('W')!.getCell('A1').value).toBe('written');
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

// ---------------------------------------------------------------------------
// atomic file target
// ---------------------------------------------------------------------------

describe('write — atomic file target', () => {
  /** Stand-in for a file the user cannot afford to lose. */
  const ORIGINAL = Buffer.from('the workbook that was already there');

  /** Scratch files the writer left behind in `dir` (the destination excluded). */
  const scratchFiles = (dir: string): string[] =>
    readdirSync(dir).filter((name) => name !== 'out.xlsx');

  it('does not touch the destination until commit()', async () => {
    const dir = tempDir();
    const path = join(dir, 'out.xlsx');
    try {
      writeFileSync(path, ORIGINAL);
      const wb = await WorkbookWriter.create({ file: path });
      // create() opens nothing: not the destination, not even a scratch file.
      expect(scratchFiles(dir)).toEqual([]);
      expect(readFileSync(path)).toEqual(ORIGINAL);

      const ws = wb.addWorksheet('S');
      ws.addRow(['a']); // still buffered: not a single byte has been produced
      expect(scratchFiles(dir)).toEqual([]);
      ws.addRow(['b']); // flushes row 1: the first ZIP bytes hit the scratch file
      expect(scratchFiles(dir)).toHaveLength(1);
      expect(scratchFiles(dir)[0]).toMatch(/^out\.xlsx\..*\.tmp$/);
      // …and the destination is still the file it was.
      expect(readFileSync(path)).toEqual(ORIGINAL);

      await wb.commit();
      expect(scratchFiles(dir)).toEqual([]);
      const ws2 = Workbook.read(new Uint8Array(readFileSync(path))).getWorksheet('S')!;
      expect(ws2.getCell('A1').value).toBe('a');
      expect(ws2.getCell('A2').value).toBe('b');
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('leaves the original byte-identical when the row loop throws', async () => {
    const dir = tempDir();
    const path = join(dir, 'out.xlsx');
    try {
      writeFileSync(path, ORIGINAL);
      const wb = await WorkbookWriter.create({ file: path });
      const ws = wb.addWorksheet('S');
      ws.addRow(['ok']);
      ws.addRow(['also ok']); // flushes row 1: bytes are on disk by now
      expect(scratchFiles(dir)).toHaveLength(1);
      try {
        ws.addRow([Number.NaN]); // the user's own loop blows up mid-write
        expect.unreachable();
      } catch (error) {
        expect((error as SheetLiteError).code).toBe('BAD_NUMBER');
      }

      wb.abort();
      expect(readFileSync(path)).toEqual(ORIGINAL);
      expect(scratchFiles(dir)).toEqual([]);
      // An aborted writer is sealed, exactly like a committed one…
      expect(wb.sealed).toBe(true);
      expect(() => ws.addRow(['late'])).toThrow(SheetLiteError);
      // …so a late commit() cannot resurrect it either.
      await wb.commit();
      expect(readFileSync(path)).toEqual(ORIGINAL);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('aborts on scope exit with `await using`, committing nothing', async () => {
    const dir = tempDir();
    const path = join(dir, 'out.xlsx');
    try {
      writeFileSync(path, ORIGINAL);
      {
        await using wb = await WorkbookWriter.create({ file: path });
        const ws = wb.addWorksheet('S');
        ws.addRow(['a']);
        ws.addRow(['b']); // a scratch file exists at this point
        expect(scratchFiles(dir)).toHaveLength(1);
        // …and no commit(): the scope simply ends.
      }
      expect(readFileSync(path)).toEqual(ORIGINAL);
      expect(scratchFiles(dir)).toEqual([]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('is a no-op once commit() has replaced the destination', async () => {
    const dir = tempDir();
    const path = join(dir, 'out.xlsx');
    try {
      writeFileSync(path, ORIGINAL);
      const wb = await WorkbookWriter.create({ file: path });
      wb.addWorksheet('S').addRow(['a']);
      await wb.commit();
      const committed = readFileSync(path);
      expect(Workbook.read(new Uint8Array(committed)).getWorksheet('S')!.getCell('A1').value).toBe('a');

      wb.abort();
      wb.abort(); // idempotent
      await wb[Symbol.asyncDispose]();
      expect(readFileSync(path)).toEqual(committed);
      expect(scratchFiles(dir)).toEqual([]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('rolls the scratch file back when commit() itself fails', async () => {
    const dir = tempDir();
    try {
      // A directory sits where the workbook should go, so the final rename
      // cannot succeed. One row only, so the first byte is produced *inside*
      // commit() and the failure lands on the rollback path.
      mkdirSync(join(dir, 'out.xlsx'));
      const wb = await WorkbookWriter.create({ file: join(dir, 'out.xlsx') });
      wb.addWorksheet('S').addRow(['a']);
      await expect(wb.commit()).rejects.toThrow();
      expect(scratchFiles(dir)).toEqual([]);
      expect(statSync(join(dir, 'out.xlsx')).isDirectory()).toBe(true);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('cannot un-send bytes to a sink: abort() only stops the writer', async () => {
    const chunks: Uint8Array[] = [];
    const wb = await WorkbookWriter.create({ sink: { write: (chunk) => chunks.push(chunk) } });
    const ws = wb.addWorksheet('S');
    ws.addRow(['a']);
    ws.addRow(['b']);
    const written = chunks.length;
    expect(written).toBeGreaterThan(0);

    wb.abort();
    expect(chunks).toHaveLength(written); // nothing more was produced
    expect(wb.sealed).toBe(true);

    // What the sink already took is a half-streamed archive it cannot give
    // back: that is the trade-off a custom target makes, documented above.
    const total = chunks.reduce((n, chunk) => n + chunk.length, 0);
    const partial = new Uint8Array(total);
    let at = 0;
    for (const chunk of chunks) {
      partial.set(chunk, at);
      at += chunk.length;
    }
    expect(() => readCentralDirectory(partial)).toThrow(/EOCD/);
  });
});

// ---------------------------------------------------------------------------
// cross-reads
// ---------------------------------------------------------------------------

describe('write — cross-reads (SheetJS / ExcelJS)', () => {
  async function sampleWorkbook(): Promise<Uint8Array> {
    return build((wb) => {
      const ws = wb.addWorksheet('Data', { views: [{ state: 'frozen', ySplit: 1 }] });
      ws.getColumn('A').width = 18;
      const header = ws.addRow(['name', 'amount', 'when', 'ok']);
      header.font = { bold: true };
      const row = ws.addRow([
        'Total',
        1234.56,
        new Date(Date.UTC(2024, 4, 17)),
        true,
      ]);
      row.getCell(2).numFmt = '#,##0.00';
      row.getCell(3).numFmt = 'yyyy-mm-dd';
      ws.addRow([]);
    });
  }

  it('SheetJS reads every written value', async () => {
    const bytes = await sampleWorkbook();
    const wb = XLSX.read(bytes, { type: 'array' });
    expect(wb.SheetNames).toEqual(['Data']);
    const sheet = wb.Sheets['Data']!;
    expect(sheet['!ref']).toBe('A1:D3');
    expect(sheet['A2']!.v).toBe('Total');
    expect(sheet['B2']!.v).toBe(1234.56);
    // The date cell carries the serial plus the `yyyy-mm-dd` format; SheetJS
    // renders it from the format (it only materialises Dates with cellDates).
    expect(sheet['C2']!.t).toBe('n');
    expect(XLSX.SSF.format('yyyy-mm-dd', sheet['C2']!.v)).toBe('2024-05-17');
    expect(sheet['D2']!.v).toBe(true);
    expect(XLSX.utils.sheet_to_json(sheet, { header: 1 })[0]).toEqual([
      'name',
      'amount',
      'when',
      'ok',
    ]);
  });

  it('ExcelJS reads values, widths and panes back', async () => {
    const bytes = await sampleWorkbook();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes as unknown as ExcelLoadInput);
    const ws = wb.getWorksheet('Data')!;
    expect(ws.getCell('A2').value).toBe('Total');
    expect(ws.getCell('B2').value).toBe(1234.56);
    expect(ws.getCell('B2').numFmt).toBe('#,##0.00');
    expect(ws.getCell('C2').value).toBeInstanceOf(Date);
    expect(ws.getCell('D2').value).toBe(true);
    expect(ws.getColumn(1).width).toBeCloseTo(18, 5);
    expect(ws.views[0]).toMatchObject({ state: 'frozen', ySplit: 1 });
    expect(ws.getCell('A1').font).toMatchObject({ bold: true });
  });

  it('re-opens with no repair (corpus-validator style)', async () => {
    const bytes = await sampleWorkbook();

    // (a) our own reader: every cell, twice (eager + streaming semantics).
    const reread = Workbook.read(bytes);
    const cells = readCells(bytes, 'Data');
    expect(cells).toHaveLength(8);
    expect(cells.map(([ref]) => ref)).toEqual([
      'A1',
      'B1',
      'C1',
      'D1',
      'A2',
      'B2',
      'C2',
      'D2',
    ]);
    expect(plain(cells[5]![1])).toBe(1234.56);

    // (b) OPC coherence: every part referenced by the workbook exists, and
    // every worksheet part is registered.
    const store = reread.parts();
    const workbookXml = new TextDecoder().decode(store.get('xl/workbook.xml')!);
    const rels = new TextDecoder().decode(store.get('xl/_rels/workbook.xml.rels')!);
    for (const rId of [...workbookXml.matchAll(/r:id="(rId\d+)"/g)].map((m) => m[1]!)) {
      expect(rels).toContain(`Id="${rId}"`);
    }
    const contentTypes = new TextDecoder().decode(store.get('[Content_Types].xml')!);
    for (const name of store.names()) {
      if (!name.endsWith('.xml') || name.startsWith('_rels') || name.includes('/_rels/')) continue;
      if (name === '[Content_Types].xml') continue;
      expect(contentTypes, `content type of ${name}`).toContain(`/${name}`);
    }

    // (c) the ZIP container itself parses with an independent walker.
    const { entries } = readCentralDirectory(bytes);
    expect(entries.every((entry) => crc32(readEntryBytes(bytes, entry)) === entry.crc)).toBe(true);

    // (d) SheetJS and ExcelJS both open it and see the same values.
    const sheetJs = XLSX.read(bytes, { type: 'array' }).Sheets['Data']!;
    expect(sheetJs['A1']!.v).toBe('name');
    const excel = new ExcelJS.Workbook();
    await excel.xlsx.load(bytes as unknown as ExcelLoadInput);
    expect(excel.getWorksheet('Data')!.getCell('A2').value).toBe('Total');

    // (e) LibreOffice, when installed: a file needing repair fails to convert.
    const soffice = ['soffice', 'libreoffice'].find((bin) => {
      const probe = spawnSync(bin, ['--version'], { stdio: 'ignore', shell: false });
      return !probe.error;
    });
    if (soffice) {
      const dir = tempDir();
      try {
        const input = join(dir, 'probe.xlsx');
        const { writeFileSync } = await import('node:fs');
        writeFileSync(input, bytes);
        const res = spawnSync(soffice, ['--headless', '--convert-to', 'xlsx', '--outdir', dir, input], {
          stdio: 'ignore',
        });
        expect(res.status, 'LibreOffice must convert the file without repair').toBe(0);
      } finally {
        rmSync(dir, { recursive: true, force: true });
      }
    }
  });
});

// ---------------------------------------------------------------------------
// bounded memory
// ---------------------------------------------------------------------------

describe('write — bounded memory', () => {
  const ROWS = 50_000;
  const COLS = 8;

  /**
   * A writer's real footprint: `TypedArray` backing stores (the deflate window,
   * the archive chunks, the encoded rows) live OUTSIDE the V8 heap, so
   * `heapUsed` alone would report a few hundred kB for a writer holding
   * megabytes. Everything below measures heap + arrayBuffers.
   */
  function memory(): number {
    const usage = process.memoryUsage();
    return usage.heapUsed + usage.arrayBuffers;
  }

  async function settle(): Promise<void> {
    const gc = (globalThis as { gc?: () => void }).gc;
    for (let i = 0; i < 3 && gc; i += 1) {
      gc();
      await new Promise((resolve) => setTimeout(resolve, 25));
    }
  }

  /** Feed `ROWS` rows through one writer and report the live-memory peak. */
  async function measure(target: { file: string } | null): Promise<number> {
    const gc = (globalThis as { gc?: () => void }).gc!;
    const wb = target
      ? await WorkbookWriter.create({ file: target.file })
      : await WorkbookWriter.create();
    const ws = wb.addWorksheet('Data');
    const baseline = memory();
    let peak = 0;
    let sinceGc = 0;
    for (let r = 1; r <= ROWS; r += 1) {
      const row = ws.addRow([
        `row-${r}`,
        r,
        r * 1.5,
        r % 2 === 0,
        `text ${r}`,
        r * 0.01,
        r,
        `padding-${'x'.repeat(40)}-${r}`,
      ]);
      row.getCell(2).numFmt = '#,##0';
      sinceGc += COLS;
      if (sinceGc >= 4096) {
        sinceGc = 0;
        gc();
        const used = memory() - baseline;
        if (used > peak) peak = used;
      }
    }
    if (target) await wb.commit();
    else await wb.toBuffer();
    gc();
    const tail = memory() - baseline;
    return Math.max(0, Math.max(peak, tail));
  }

  it(
    'streams 50k rows to a file with a bounded live heap',
    { timeout: 300_000 },
    async () => {
      const dir = tempDir();
      const path = join(dir, 'big.xlsx');
      try {
        await settle();
        expect(typeof (globalThis as { gc?: () => void }).gc).toBe('function'); // --expose-gc

        const streamed = await measure({ file: path });
        const bytes = new Uint8Array(readFileSync(path));
        const { entries } = readCentralDirectory(bytes);
        const sheet = entries.find((entry) => entry.name === 'xl/worksheets/sheet1.xml')!;
        expect(sheet.uncompressedSize).toBeGreaterThan(4 * 1024 * 1024);

        // Same rows, buffered target: the whole archive is retained instead.
        await settle();
        const buffered = await measure(null);

        const mib = (bytesCount: number): string => `${(bytesCount / 1024 / 1024).toFixed(1)} MB`;
        console.log(
          `[memory] write ${ROWS} rows x ${COLS} cols — file target ${mib(streamed)} vs buffer target ${mib(buffered)}`,
        );

        // The file target keeps only the deflate window, one row and the ZIP
        // metadata: 32 MB is a very generous ceiling for it…
        expect(streamed).toBeLessThan(32 * 1024 * 1024);
        // …and the produced sheet is an order of magnitude bigger than that,
        // so the bound cannot come from a small sheet.
        expect(sheet.uncompressedSize).toBeGreaterThan(streamed * 2);
        // Buffering the whole archive is strictly more expensive.
        expect(buffered).toBeGreaterThan(streamed);

        // …and the file is a real 50 000-row workbook.
        const { StreamingWorkbook } = await import('../src/stream');
        const reader = await StreamingWorkbook.open(path);
        try {
          let rows = 0;
          for await (const row of reader.readRows('Data')) {
            if (rows === 0) expect(row.cells[0]!.value).toBe('row-1');
            rows += 1;
          }
          expect(rows).toBe(ROWS);
        } finally {
          await reader.close();
        }
      } finally {
        rmSync(dir, { recursive: true, force: true });
      }
    },
  );
});

// ---------------------------------------------------------------------------
// value helpers re-exported for the type surface
// ---------------------------------------------------------------------------

describe('write — value type surface', () => {
  it('accepts the same CellValue union the reader produces', async () => {
    const values: CellValue[] = [
      null,
      1,
      'a',
      true,
      new Date(Date.UTC(2024, 0, 2)),
      { error: '#N/A' },
      { richText: [{ text: 'r' }] },
      { formula: 'A1+1', result: 2 },
    ];
    const bytes = await build((wb) => {
      const row = wb.addWorksheet('V').addRow(values);
      // A Date only reads back as a Date when the cell carries a date format
      // (the same rule the reader applies).
      row.getCell(5).numFmt = 'yyyy-mm-dd';
    });
    const cells = readCells(bytes, 'V');
    // `null` produces no cell at all, exactly like the patch writer.
    expect(cells).toHaveLength(values.length - 1);
    expect(cells.map(([ref]) => ref)).toEqual(['B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1']);
    expect(cells.map(([, value]) => plain(value))).toEqual([
      1,
      'a',
      true,
      '2024-01-02T00:00:00.000Z',
      { error: '#N/A' },
      'r',
      { formula: 'A1+1' },
    ]);
  });
});
