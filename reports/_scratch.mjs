import { zipSync, strToU8, strFromU8 } from 'fflate';
import { Workbook, serialToDate, dateToSerial, parseXml, refToCoords, isDateFormatCode, isSafeEntryName } from '../dist/index.mjs';

const decodeEntities = (v) => parseXml(`<x>${v}</x>`).children[0].text;

const line = (s) => console.log('\n=== ' + s + ' ===');

line('1. 1900 leap-year / pre-1900-03-01 mapping');
for (const s of [1, 59, 60, 61, 25569]) {
  console.log(`serial ${s} ->`, serialToDate(s, false).toISOString());
}
console.log('dateToSerial(1900-01-01) ->', dateToSerial(new Date(Date.UTC(1900, 0, 1)), false));
console.log('dateToSerial(1900-02-28) ->', dateToSerial(new Date(Date.UTC(1900, 1, 28)), false));
console.log('dateToSerial(1900-03-01) ->', dateToSerial(new Date(Date.UTC(1900, 2, 1)), false));

line('2. entity decoding of out-of-range code points (via parseXml)');
for (const e of ['&amp;', '&#65;', '&#x1F600;', '&#x110000;', '&#0;', '&#xD800;']) {
  try { console.log(e, '->', JSON.stringify(parseXml(`<a b="${e}"/>`).children[0].attrs.b)); }
  catch (err) { console.log(e, 'THREW', err.constructor.name, err.message); }
}
try { parseXml('<a><t>&#x110000;</t></a>'); console.log('parseXml text out-of-range: returned (decoded later)'); }
catch (err) { console.log('parseXml text out-of-range THREW', err.constructor.name, err.message); }

line('3. attribute value containing ">"');
const px = parseXml('<a b="x>y"><c/></a>');
console.log('root children:', px.children.map((n) => n.name));
console.log('a attrs:', JSON.stringify(px.children[0]?.attrs));
console.log('a text:', JSON.stringify(px.children[0]?.text));

line('4. refToCoords on malformed refs');
for (const r of ['A1', 'aa10', 'foo', 'A0', '', 'A1:B2']) console.log(JSON.stringify(r), '->', JSON.stringify(refToCoords(r)));

function minimalXlsx(sheetXml) {
  const ct = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`;
  const rels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const wb = `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const wbrels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;
  return zipSync({
    '[Content_Types].xml': strToU8(ct),
    '_rels/.rels': strToU8(rels),
    'xl/workbook.xml': strToU8(wb),
    'xl/_rels/workbook.xml.rels': strToU8(wbrels),
    'xl/worksheets/sheet1.xml': strToU8(sheetXml),
  });
}

const sheetOf = (buf) => strFromU8(Workbook.read(buf).parts().get('xl/worksheets/sheet1.xml'));

line('5. cell without r attribute (positional cells)');
{
  const x = minimalXlsx(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c><v>5</v></c><c r="B1"><v>9</v></c></row></sheetData></worksheet>`);
  const ws = Workbook.read(x).getWorksheet('S');
  console.log('A1 =', ws.getCell('A1').value, '| B1 =', ws.getCell('B1').value, '| rowCount =', ws.rowCount);
  const w = Workbook.read(x); w.getWorksheet('S').getCell('A1').value = 111;
  const out = w.toBuffer();
  console.log('EDITED sheet:', sheetOf(out));
  const rr = Workbook.read(out).getWorksheet('S');
  console.log('reread A1 =', rr.getCell('A1').value, 'A2 =', rr.getCell('A2').value);
}

line('6. inserting a cell into an existing self-closing row');
{
  const x = minimalXlsx(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"/><row r="2"><c r="A2"><v>5</v></c></row></sheetData></worksheet>`);
  const w = Workbook.read(x);
  w.getWorksheet('S').getCell('A1').value = 42;
  const out = w.toBuffer();
  console.log('EDITED sheet:', sheetOf(out));
  const rr = Workbook.read(out).getWorksheet('S');
  console.log('reread A1 =', rr.getCell('A1').value, 'A2 =', rr.getCell('A2').value);
}

line('7. formula without cached result round-trip');
{
  const x = minimalXlsx(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>1+1</f></c></row></sheetData></worksheet>`);
  const w = Workbook.read(x);
  console.log('read formula w/o v:', JSON.stringify(w.getWorksheet('S').getCell('A1').value));
  w.getWorksheet('S').getCell('B1').value = { formula: '2+2' };
  const out = w.toBuffer();
  console.log('sheet after edit:', sheetOf(out));
  console.log('reread B1 =', JSON.stringify(Workbook.read(out).getWorksheet('S').getCell('B1').value));
}

line('8. empty <v></v> numeric cell');
{
  const x = minimalXlsx(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v></v></c></row></sheetData></worksheet>`);
  console.log('A1 =', JSON.stringify(Workbook.read(x).getWorksheet('S').getCell('A1').value));
}

line('9. shared-string index out of range / empty v');
{
  const x = minimalXlsx(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`);
  try { console.log('A1 =', JSON.stringify(Workbook.read(x).getWorksheet('S').getCell('A1').value)); }
  catch (e) { console.log('threw', e.message); }
}

line('10. isDateFormatCode on common formats');
for (const c of ['yyyy-mm-dd', '[$-409]d-mmm-yy', '[$-F800]dddd, mmmm dd, yyyy', '[h]:mm:ss', '0.00', '#,##0.00\\ "€"', 'dd/mm/yyyy;@', 'General']) {
  console.log(JSON.stringify(c), '->', isDateFormatCode(c));
}

line('11. shared-string cell edit converts to inlineStr');
{
  const ss = `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>shared</t></si></sst>`;
  const sheet = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`;
  const ct = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`;
  const rels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const wb = `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const wbrels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`;
  const x = zipSync({ '[Content_Types].xml': strToU8(ct), '_rels/.rels': strToU8(rels), 'xl/workbook.xml': strToU8(wb), 'xl/_rels/workbook.xml.rels': strToU8(wbrels), 'xl/worksheets/sheet1.xml': strToU8(sheet), 'xl/sharedStrings.xml': strToU8(ss) });
  const w = Workbook.read(x);
  console.log('read A1 =', JSON.stringify(w.getWorksheet('S').getCell('A1').value));
  w.getWorksheet('S').getCell('A1').value = 'changed';
  const out = w.toBuffer();
  console.log('EDITED sheet:', sheetOf(out));
  console.log('sharedStrings after edit:', strFromU8(Workbook.read(out).parts().get('xl/sharedStrings.xml')));
  console.log('reread A1 =', JSON.stringify(Workbook.read(out).getWorksheet('S').getCell('A1').value));
}

line('12. entity / emoji / whitespace round-trip through edit');
{
  const vals = ['a & b < c > d " e \' f', '😀🎉', '  leading', 'trailing  ', 'tab\there', 'new\nline', '<t>x</t>', '&amp;'];
  for (const v of vals) {
    const x = minimalXlsx(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"></row></sheetData></worksheet>`);
    const w = Workbook.read(x);
    w.getWorksheet('S').getCell('A1').value = v;
    const back = Workbook.read(w.toBuffer()).getWorksheet('S').getCell('A1').value;
    console.log(JSON.stringify(v), '->', JSON.stringify(back), back === v ? 'OK' : '*** MISMATCH');
  }
}

line('13. inserting a new cell between existing cells (column order)');
{
  const x = minimalXlsx(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row></sheetData></worksheet>`);
  const w = Workbook.read(x);
  w.getWorksheet('S').getCell('B1').value = 2;
  console.log('EDITED sheet:', sheetOf(w.toBuffer()));
}

line('14. addRow on a sheet whose sheetData is self-closing with trailing elements');
{
  const x = minimalXlsx(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);
  const w = Workbook.read(x);
  w.getWorksheet('S').addRow(['x']);
  console.log('EDITED sheet:', sheetOf(w.toBuffer()));
}

line('15. ZIP traversal variants');
for (const n of ['..\\evil.txt', 'a\\..\\b', 'C:\\x', '..%2fevil', '....//evil', 'a/b/../../../c', '\\evil', 'foo/..']) {
  console.log(JSON.stringify(n), '->', isSafeEntryName(n));
}
line('16. ZIP side note: limits are enforced after full decompression by unzipSync');
console.log('readZip source uses unzipSync(raw) with no filter; entry/size/ratio checks run on the already-decompressed map.');

