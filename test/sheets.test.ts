import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { strFromU8, strToU8, zipSync } from 'fflate';
import { Workbook, SheetLiteError } from '../src/index';
import { parseContentTypes, contentTypeOf, resolvePartPath } from '../src/index';
import { parseXml } from '../src/xml/xml';
import { makeXlsxWithSheet } from './fixtures/make';

const eq = (a: Uint8Array, b: Uint8Array): boolean => Buffer.from(a).equals(Buffer.from(b));

const WORKSHEET_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

/** Build a realistic multi-sheet workbook with ExcelJS (in-test fixture). */
async function buildExcelJsFile(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  data.getCell('A1').value = 'Hello';
  data.getCell('A1').font = { bold: true };
  data.getCell('A2').value = 42;
  data.getCell('A3').value = 1.5;
  data.getCell('A4').value = new Date(Date.UTC(2026, 0, 15));
  data.getCell('A4').numFmt = 'yyyy-mm-dd';
  data.getCell('B1').value = { formula: 'SUM(A2:A3)', result: 43.5 };
  data.getCell('B2').value = 'keep me';
  data.getCell('C1').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"red,green,blue"'],
  };
  data.mergeCells('D1:E1');
  data.getCell('D1').value = 'merged';
  // Give the Data sheet an existing sheetView with a pane so pane *replacement*
  // can be tested (ExcelJS omits sheetViews entirely when no views are set).
  data.views = [{ state: 'frozen', xSplit: 1 }];

  const config = wb.addWorksheet('Config');
  config.getCell('A1').value = 'key';
  config.getCell('B1').value = 'value';

  const notes = wb.addWorksheet('Notes');
  notes.getCell('A1').value = 'note';

  return Buffer.from(await wb.xlsx.writeBuffer());
}

/** Two-sheet package without styles/sharedStrings plus a calcChain,
 *  defined names and bookViews — for the package-level edge cases. */
function makeTwoSheetPackage(): Uint8Array {
  const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="${WORKSHEET_CT}"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="${WORKSHEET_CT}"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView activeTab="1"/></bookViews>
  <sheets>
    <sheet name="One" sheetId="1" r:id="rId1"/>
    <sheet name="Two" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <definedName name="OnesRange">One!$A$1</definedName>
    <definedName name="TwosRange" localSheetId="1">Two!$A$1</definedName>
  </definedNames>
  <calcPr calcId="171027"/>
</workbook>`;
  const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>`;
  const sheet = (n: number) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>${n}</v></c></row></sheetData>
</worksheet>`;
  const calcChain = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>`;
  return zipSync({
    '[Content_Types].xml': strToU8(ct),
    '_rels/.rels': strToU8(rootRels),
    'xl/workbook.xml': strToU8(workbook),
    'xl/_rels/workbook.xml.rels': strToU8(wbRels),
    'xl/worksheets/sheet1.xml': strToU8(sheet(1)),
    'xl/worksheets/sheet2.xml': strToU8(sheet(2)),
    'xl/calcChain.xml': strToU8(calcChain),
  });
}

const allElements = (name: string, node: ReturnType<typeof parseXml>): ReturnType<typeof parseXml>[] => {
  const out: ReturnType<typeof parseXml>[] = [];
  const walk = (n: ReturnType<typeof parseXml>): void => {
    for (const child of n.children) {
      if (child.name === name) out.push(child);
      walk(child);
    }
  };
  walk(node);
  return out;
};

/**
 * Package coherence: every Override points at an existing part, every
 * relationship target exists, every r:id used by workbook.xml resolves, and
 * sheet ids/names/relationship ids are unique.
 */
function expectCoherentPackage(bytes: Uint8Array): void {
  const wb = Workbook.read(bytes);
  const parts = wb.parts();

  const ct = parseContentTypes(parts.get('[Content_Types].xml')!);
  for (const partName of ct.overrides.keys()) {
    expect(parts.has(partName.replace(/^\//, '')), `override without part: ${partName}`).toBe(true);
  }

  const relTargets = new Map<string, string>();
  for (const rel of allElements('Relationship', parseXml(strFromU8(parts.get('xl/_rels/workbook.xml.rels')!)))) {
    const id = rel.attrs.Id ?? '';
    expect(relTargets.has(id), `duplicate relationship id: ${id}`).toBe(false);
    const resolved = resolvePartPath('/xl/workbook.xml', rel.attrs.Target ?? '');
    relTargets.set(id, resolved);
    expect(parts.has(resolved.replace(/^\//, '')), `relationship target missing: ${resolved}`).toBe(true);
  }

  const sheetIds = new Set<string>();
  const names = new Set<string>();
  for (const sheet of allElements('sheet', parseXml(strFromU8(parts.get('xl/workbook.xml')!)))) {
    const rId = sheet.attrs['r:id'] ?? '';
    expect(relTargets.has(rId), `sheet r:id without a relationship: ${rId}`).toBe(true);
    expect(sheet.attrs.sheetId, 'sheet must carry a sheetId').toBeTruthy();
    sheetIds.add(sheet.attrs.sheetId!);
    names.add(sheet.attrs.name ?? '');
  }
  expect(sheetIds.size).toBe(wb.sheetNames.length);
  expect(names.size).toBe(wb.sheetNames.length);
  expect(wb.metadata.sheets.length).toBe(wb.sheetNames.length);
}

describe('sheets — addWorksheet', () => {
  it('creates a sheet on a real ExcelJS workbook that survives a round-trip', async () => {
    const file = await buildExcelJsFile();
    const wb = Workbook.read(file);
    const ws = wb.addWorksheet('Summary');

    expect(wb.sheetNames).toEqual(['Data', 'Config', 'Notes', 'Summary']);
    expect(wb.worksheets).toContain(ws);
    expect(wb.getWorksheet('Summary')).toBe(ws);
    expect(wb.getWorksheet(ws.id)).toBe(ws);

    const out = wb.toBuffer();
    expect(wb.parts().has('xl/worksheets/sheet4.xml')).toBe(true);

    const reread = Workbook.read(out);
    expect(reread.sheetNames).toEqual(['Data', 'Config', 'Notes', 'Summary']);
    const fresh = reread.getWorksheet('Summary')!;
    expect(fresh).toBeDefined();
    expect(fresh.getCell('A1').value).toBeNull();
    expect(contentTypeOf(reread.contentTypes(), '/xl/worksheets/sheet4.xml')).toBe(WORKSHEET_CT);
    expectCoherentPackage(out);
  });

  it('values written into the new sheet round-trip, keep styles and are visible to SheetJS', async () => {
    const file = await buildExcelJsFile();
    const wb = Workbook.read(file);
    const ws = wb.addWorksheet('Summary');
    ws.getCell('A1').value = 'Hello summary';
    ws.getCell('B2').value = 7.5;
    ws.getCell('A3').value = true;
    ws.getCell('A4').value = { formula: 'B2*2', result: 15 };
    const date = new Date(Date.UTC(2026, 5, 30));
    ws.getCell('A5').value = date;
    ws.getCell('A5').numFmt = 'yyyy-mm-dd';

    const out = wb.toBuffer();
    const reread = Workbook.read(out).getWorksheet('Summary')!;
    expect(reread.getCell('A1').value).toBe('Hello summary');
    expect(reread.getCell('B2').value).toBe(7.5);
    expect(reread.getCell('A3').value).toBe(true);
    expect(reread.getCell('A4').value).toEqual({ formula: 'B2*2', result: 15 });
    expect((reread.getCell('A5').value as Date).toISOString()).toBe('2026-06-30T00:00:00.000Z');
    expect(reread.getCell('A5').numFmt).toBe('yyyy-mm-dd');

    // SheetJS still parses the result and sees the new values (dimension kept in sync).
    const sjs = XLSX.read(out, { type: 'buffer', cellDates: true });
    expect(sjs.SheetNames).toContain('Summary');
    const sheet = sjs.Sheets.Summary!;
    expect(sheet.A1.v).toBe('Hello summary');
    expect(sheet.B2.v).toBe(7.5);
    expect(sheet.A5.v).toBeInstanceOf(Date);
    expect(sheet['!ref']).toBe('A1:B5');
    expectCoherentPackage(out);
  });

  it('works when styles.xml and sharedStrings.xml are missing', () => {
    const wb = Workbook.read(makeTwoSheetPackage());
    expect(wb.sheetNames).toEqual(['One', 'Two']);
    const ws = wb.addWorksheet('Fresh');
    ws.getCell('A1').value = 'inline string';
    ws.getCell('B1').value = 3;

    const out = wb.toBuffer();
    const reread = Workbook.read(out);
    expect(reread.sheetNames).toEqual(['One', 'Two', 'Fresh']);
    expect(reread.parts().has('xl/styles.xml')).toBe(true);
    const fresh = reread.getWorksheet('Fresh')!;
    expect(fresh.getCell('A1').value).toBe('inline string');
    expect(fresh.getCell('B1').value).toBe(3);
    expect(fresh.getCell('A1').numFmt).toBe('General');

    const sjs = XLSX.read(out, { type: 'buffer' });
    expect(sjs.SheetNames).toEqual(['One', 'Two', 'Fresh']);
    expect(sjs.Sheets.Fresh!.A1.v).toBe('inline string');
    expectCoherentPackage(out);
  });

  it('rejects invalid names with coded errors', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    expect(() => wb.addWorksheet('Data')).toThrow(SheetLiteError);
    try {
      wb.addWorksheet('Data');
      expect.unreachable();
    } catch (error) {
      expect((error as SheetLiteError).code).toBe('SHEET_EXISTS');
    }
    for (const bad of ['', 'a'.repeat(32), 'has:colon', 'back\\slash', 'sl/ash', 'que?ry', 'st*ar', 'br[ack']) {
      try {
        wb.addWorksheet(bad);
        expect.unreachable();
      } catch (error) {
        expect((error as SheetLiteError).code).toBe('BAD_SHEET_NAME');
      }
    }
  });

  it('reads views written by ExcelJS back as frozen panes', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const views = wb.getWorksheet('Data')!.views;
    expect(views[0]).toMatchObject({ state: 'frozen', xSplit: 1 });
  });

  it('supports frozen panes at creation time', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const ws = wb.addWorksheet('Frozen', { views: [{ state: 'frozen', xSplit: 1, ySplit: 2 }] });

    const reread = Workbook.read(wb.toBuffer()).getWorksheet('Frozen')!;
    const view = reread.views[0]!;
    expect(view.state).toBe('frozen');
    expect(view.xSplit).toBe(1);
    expect(view.ySplit).toBe(2);
    expect(view.topLeftCell).toBe('B3');
    expect(ws.views[0]!.state).toBe('frozen');
  });
});

describe('sheets — removeWorksheet', () => {
  it('removes a middle sheet and keeps the other sheets intact', async () => {
    const file = await buildExcelJsFile();
    const wb = Workbook.read(file);
    wb.removeWorksheet('Config');

    expect(wb.sheetNames).toEqual(['Data', 'Notes']);
    const out = wb.toBuffer();

    const reread = Workbook.read(out);
    expect(reread.sheetNames).toEqual(['Data', 'Notes']);
    expect(reread.getWorksheet('Config')).toBeUndefined();
    expect(reread.parts().has('xl/worksheets/sheet2.xml')).toBe(false);

    const data = reread.getWorksheet('Data')!;
    expect(data.getCell('A1').value).toBe('Hello');
    expect(data.getCell('B1').value).toEqual({ formula: 'SUM(A2:A3)', result: 43.5 });
    expect(data.getCell('D1').value).toBe('merged');
    expect(strFromU8(reread.parts().get('xl/worksheets/sheet1.xml')!)).toContain('<mergeCells');
    expect(reread.getWorksheet('Notes')!.getCell('A1').value).toBe('note');

    const sjs = XLSX.read(out, { type: 'buffer' });
    expect(sjs.SheetNames).toEqual(['Data', 'Notes']);
    expectCoherentPackage(out);
  });

  it('leaves every other part byte-identical when removing a sheet', async () => {
    const file = await buildExcelJsFile();
    const wb = Workbook.read(file);
    wb.removeWorksheet('Config');
    const out = wb.toBuffer();

    const before = Workbook.read(file).parts();
    const after = Workbook.read(out).parts();
    const changed = new Set<string>();
    for (const name of before.names()) {
      if (!after.has(name)) {
        changed.add(name);
        continue;
      }
      if (!eq(before.get(name)!, after.get(name)!)) changed.add(name);
    }
    expect(changed).toEqual(
      new Set(['[Content_Types].xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet2.xml']),
    );
    for (const name of after.names()) {
      if (!before.has(name)) expect(changed.has(name), `new part not accounted: ${name}`).toBe(true);
    }
  });

  it('removes a freshly added sheet leaving every part content-identical (patch-only proof)', async () => {
    const file = await buildExcelJsFile();
    const wb = Workbook.read(file);
    wb.addWorksheet('Temp');
    wb.removeWorksheet('Temp');
    const out = wb.toBuffer();

    const before = Workbook.read(file).parts();
    const after = Workbook.read(out).parts();
    expect(after.names().sort()).toEqual(before.names().sort());
    for (const name of before.names()) {
      expect(eq(before.get(name)!, after.get(name)!), `part changed: ${name}`).toBe(true);
    }
  });

  it('drops a stale calcChain with its relationship and override, fixes names and activeTab', () => {
    const bytes = makeTwoSheetPackage();
    const wb = Workbook.read(bytes);
    wb.removeWorksheet('One');
    const out = wb.toBuffer();

    const reread = Workbook.read(out);
    expect(reread.sheetNames).toEqual(['Two']);
    expect(reread.parts().has('xl/calcChain.xml')).toBe(false);

    const relsXml = strFromU8(reread.parts().get('xl/_rels/workbook.xml.rels')!);
    expect(relsXml).not.toContain('calcChain');
    const ct = reread.contentTypes();
    expect(ct.overrides.has('/xl/calcChain.xml')).toBe(false);

    const workbookXml = strFromU8(reread.parts().get('xl/workbook.xml')!);
    expect(workbookXml).not.toContain('OnesRange');
    expect(workbookXml).toContain('TwosRange');
    expect(workbookXml).toMatch(/localSheetId="0"/);
    expect(workbookXml).toMatch(/activeTab="0"/);

    expect(reread.getWorksheet('Two')!.getCell('A1').value).toBe(2);
    const sjs = XLSX.read(out, { type: 'buffer' });
    expect(sjs.SheetNames).toEqual(['Two']);
    expectCoherentPackage(out);
  });

  it('rejects unknown sheets and the last remaining sheet with coded errors', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    try {
      wb.removeWorksheet('Nope');
      expect.unreachable();
    } catch (error) {
      expect((error as SheetLiteError).code).toBe('SHEET_NOT_FOUND');
    }
    try {
      wb.removeWorksheet(999);
      expect.unreachable();
    } catch (error) {
      expect((error as SheetLiteError).code).toBe('SHEET_NOT_FOUND');
    }

    const single = Workbook.read(makeXlsxWithSheet(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData/></worksheet>`,
    ));
    try {
      single.removeWorksheet('Sheet1');
      expect.unreachable();
    } catch (error) {
      expect((error as SheetLiteError).code).toBe('LAST_SHEET');
    }
    // The failed removal left the workbook untouched.
    expect(eq(single.toBuffer(), makeXlsxWithSheet(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData/></worksheet>`,
    ))).toBe(true);
  });
});

describe('sheets — column widths & hidden columns', () => {
  it('sets and reads back width/hidden on a new sheet (inline strings, no sharedStrings)', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const ws = wb.addWorksheet('Cols');
    ws.getColumn(1).width = 12.5;
    ws.getColumn(1).hidden = true;
    ws.getColumn('B').width = 3.5;
    expect(ws.getColumn('A').letter).toBe('A');

    const out = wb.toBuffer();
    const outParts = Workbook.read(out).parts();
    expect(outParts.has('xl/worksheets/sheet4.xml')).toBe(true);
    const sheetXml = strFromU8(outParts.get('xl/worksheets/sheet4.xml')!);
    expect(sheetXml).toContain('<cols>');
    expect(sheetXml).toContain('width="12.5" customWidth="1"');
    expect(sheetXml).toContain('hidden="1"');

    const reread = Workbook.read(out).getWorksheet('Cols')!;
    expect(reread.getColumn('A').width).toBe(12.5);
    expect(reread.getColumn('A').hidden).toBe(true);
    expect(reread.getColumn(2).width).toBe(3.5);
    expect(reread.getColumn(2).hidden).toBe(false);
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('merges into existing <col> ranges without destroying other columns', () => {
    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <cols>
    <col min="1" max="3" width="8" customWidth="1"/>
    <col min="5" max="6" width="20" customWidth="1"/>
  </cols>
  <sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>
</worksheet>`;
    const wb = Workbook.read(makeXlsxWithSheet(sheetXml));
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getColumn(2).width = 12.5; // splits the 1-3 range
    ws.getColumn(4).width = 5; // new entry between the two ranges
    ws.getColumn(6).hidden = true; // splits the 5-6 range

    const out = wb.toBuffer();
    const reread = Workbook.read(out).getWorksheet('Sheet1')!;
    expect(reread.getColumn(1).width).toBe(8);
    expect(reread.getColumn(2).width).toBe(12.5);
    expect(reread.getColumn(3).width).toBe(8);
    expect(reread.getColumn(4).width).toBe(5);
    expect(reread.getColumn(5).width).toBe(20);
    expect(reread.getColumn(6).width).toBe(20);
    expect(reread.getColumn(6).hidden).toBe(true);
    expect(reread.getColumn(5).hidden).toBe(false);
    // cells untouched
    expect(reread.getCell('A1').value).toBe(1);
    expect(reread.getCell('B1').value).toBe(2);
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('clears a width back to the sheet default', () => {
    const wb = Workbook.read(makeXlsxWithSheet(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData/></worksheet>`,
    ));
    const ws = wb.getWorksheet('Sheet1')!;
    const col = ws.getColumn(1);
    col.width = 30;
    expect(wb.toBuffer()).toBeTruthy();
    expect(Workbook.read(wb.toBuffer()).getWorksheet('Sheet1')!.getColumn(1).width).toBe(30);
    col.width = undefined;
    const reread = Workbook.read(wb.toBuffer()).getWorksheet('Sheet1')!;
    expect(reread.getColumn(1).width).toBeUndefined();
  });

  it('rejects invalid column numbers and row heights', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const ws = wb.getWorksheet('Data')!;
    expect(() => ws.getColumn(0)).toThrow(SheetLiteError);
    expect(() => ws.getColumn(-1)).toThrow(SheetLiteError);
    expect(() => ws.getColumn(16385)).toThrow(SheetLiteError);
    expect(() => ws.getColumn('ZZZ')).toThrow(SheetLiteError);
    expect(() => ws.getColumn('!!')).toThrow(SheetLiteError);
    try {
      ws.getColumn(0);
      expect.unreachable();
    } catch (error) {
      expect((error as SheetLiteError).code).toBe('BAD_COLUMN');
    }
    expect(() => {
      ws.getRow(1).height = -5;
    }).toThrow(SheetLiteError);
    try {
      ws.getRow(1).height = -5;
      expect.unreachable();
    } catch (error) {
      expect((error as SheetLiteError).code).toBe('BAD_ROW_HEIGHT');
    }
  });
});

describe('sheets — row heights', () => {
  it('patches the height of an existing row and preserves its cells', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const ws = wb.getWorksheet('Data')!;
    ws.getRow(2).height = 24.5;

    const out = wb.toBuffer();
    const sheetXml = strFromU8(Workbook.read(out).parts().get('xl/worksheets/sheet1.xml')!);
    expect(sheetXml).toMatch(/ht="24\.5" customHeight="1"/);

    const reread = Workbook.read(out).getWorksheet('Data')!;
    expect(reread.getRow(2).height).toBe(24.5);
    expect(reread.getCell('A2').value).toBe(42);
    expect(reread.getCell('B2').value).toBe('keep me');
    expect(reread.getRow(1).height).toBeUndefined();
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('creates height-only rows and combines height with new cells', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const ws = wb.getWorksheet('Data')!;
    ws.getRow(10).height = 30;
    ws.getCell('A5').value = 'new cell';
    ws.getRow(5).height = 18;

    const out = wb.toBuffer();
    const reread = Workbook.read(out).getWorksheet('Data')!;
    expect(reread.getRow(10).height).toBe(30);
    expect(reread.getRow(5).height).toBe(18);
    expect(reread.getCell('A5').value).toBe('new cell');
    expect(reread.getCell('A2').value).toBe(42);
    const sheetXml = strFromU8(Workbook.read(out).parts().get('xl/worksheets/sheet1.xml')!);
    expect(sheetXml).toContain('<row r="10" ht="30" customHeight="1"/>');
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('clears a row height (height = undefined)', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const ws = wb.getWorksheet('Data')!;
    ws.getRow(2).height = 24.5;
    ws.getRow(2).height = undefined;

    const reread = Workbook.read(wb.toBuffer()).getWorksheet('Data')!;
    expect(reread.getRow(2).height).toBeUndefined();
    expect(reread.getCell('A2').value).toBe(42);
  });
});

describe('sheets — frozen panes (views)', () => {
  it('creates sheetViews on a sheet that has none', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const ws = wb.getWorksheet('Notes')!;
    ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3 }];

    const out = wb.toBuffer();
    const reread = Workbook.read(out).getWorksheet('Notes')!;
    const view = reread.views[0]!;
    expect(view.state).toBe('frozen');
    expect(view.xSplit).toBe(2);
    expect(view.ySplit).toBe(3);
    expect(view.topLeftCell).toBe('C4');
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('freezes columns only and computes topLeftCell', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const ws = wb.getWorksheet('Notes')!;
    ws.views = [{ state: 'frozen', xSplit: 2 }];

    const reread = Workbook.read(wb.toBuffer()).getWorksheet('Notes')!;
    expect(reread.views[0]).toMatchObject({ state: 'frozen', xSplit: 2, topLeftCell: 'C1' });
    const sheetXml = strFromU8(Workbook.read(wb.toBuffer()).parts().get('xl/worksheets/sheet3.xml')!);
    expect(sheetXml).toContain('activePane="topRight"');
  });

  it('replaces an existing pane while preserving the other sheetView content', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const ws = wb.getWorksheet('Data')!;
    // The fixture Data sheet has an ExcelJS-written sheetView with a pane + selection.
    const before = strFromU8(wb.parts().get('xl/worksheets/sheet1.xml')!);
    expect(before).toContain('<pane');
    expect(before).toContain('<selection');

    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
    const out = wb.toBuffer();
    const sheetXml = strFromU8(Workbook.read(out).parts().get('xl/worksheets/sheet1.xml')!);
    expect(sheetXml).toContain('<selection');
    expect(sheetXml).toContain('<pane ySplit="1"');
    // exactly one pane, the old one replaced
    expect(sheetXml.match(/<pane /g)).toHaveLength(1);

    const reread = Workbook.read(out).getWorksheet('Data')!;
    expect(reread.views[0]).toMatchObject({ state: 'frozen', ySplit: 1 });
  });

  it('unfreezes panes (views = [])', async () => {
    const wb = Workbook.read(await buildExcelJsFile());
    const ws = wb.getWorksheet('Data')!;
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
    let reread = Workbook.read(wb.toBuffer()).getWorksheet('Data')!;
    expect(reread.views[0]!.state).toBe('frozen');

    ws.views = [];
    reread = Workbook.read(wb.toBuffer()).getWorksheet('Data')!;
    const sheetXml = strFromU8(Workbook.read(wb.toBuffer()).parts().get('xl/worksheets/sheet1.xml')!);
    expect(sheetXml).not.toContain('<pane');
    expect(reread.views[0]!.state).toBe('normal');

    // Setting a normal view is also a no-pane state.
    ws.views = [{ state: 'normal' }];
    reread = Workbook.read(wb.toBuffer()).getWorksheet('Data')!;
    expect(reread.views[0]!.state).toBe('normal');
  });
});

describe('sheets — preservation invariants', () => {
  it('keeps every unrelated part byte-identical across sheet-structure edits', async () => {
    const file = await buildExcelJsFile();
    const wb = Workbook.read(file);

    // structure: add a sheet, edit it
    const fresh = wb.addWorksheet('Summary');
    fresh.getCell('A1').value = 'probe';
    fresh.getColumn('A').width = 20;
    fresh.getRow(1).height = 22;
    fresh.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];

    // edit one existing sheet (widths/height/views/cell)
    const data = wb.getWorksheet('Data')!;
    data.getCell('A2').value = 100;
    data.getColumn(2).width = 25;
    data.getRow(3).height = 40;
    data.views = [{ state: 'frozen', xSplit: 0, ySplit: 2 }];

    const out = wb.toBuffer();
    const before = Workbook.read(file).parts();
    const after = Workbook.read(out).parts();

    // Only the package parts that register sheets, the two touched sheet parts
    // and (because the new sheet carries no styles) nothing else may change.
    const allowed = new Set([
      '[Content_Types].xml',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/worksheets/sheet1.xml',
      'xl/worksheets/sheet4.xml',
    ]);
    const changed = new Set<string>();
    for (const name of before.names()) {
      if (!after.has(name) || !eq(before.get(name)!, after.get(name)!)) changed.add(name);
    }
    for (const name of after.names()) {
      if (!before.has(name)) changed.add(name);
    }
    expect(changed).toEqual(allowed);

    expectCoherentPackage(out);
    const sjs = XLSX.read(out, { type: 'buffer' });
    expect(sjs.SheetNames).toEqual(['Data', 'Config', 'Notes', 'Summary']);
    expect(sjs.Sheets.Summary!.A1.v).toBe('probe');
  });

  it('add + edit + remove + remove keeps the remaining sheets fully intact', async () => {
    const file = await buildExcelJsFile();
    const wb = Workbook.read(file);
    const fresh = wb.addWorksheet('Temp');
    fresh.getCell('A1').value = 'temp';
    wb.removeWorksheet('Temp');
    wb.removeWorksheet('Notes');
    const out = wb.toBuffer();

    const before = Workbook.read(file).parts();
    const after = Workbook.read(out).parts();
    // Temp was added and removed again (net zero); Notes is sheet3.
    const allowed = new Set([
      '[Content_Types].xml',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/worksheets/sheet3.xml',
    ]);
    const changed = new Set<string>();
    for (const name of before.names()) {
      if (!after.has(name) || !eq(before.get(name)!, after.get(name)!)) changed.add(name);
    }
    for (const name of after.names()) {
      if (!before.has(name)) changed.add(name);
    }
    expect(changed).toEqual(allowed);

    const reread = Workbook.read(out);
    expect(reread.sheetNames).toEqual(['Data', 'Config']);
    expect(reread.getWorksheet('Data')!.getCell('A1').value).toBe('Hello');
    expectCoherentPackage(out);
  });
});