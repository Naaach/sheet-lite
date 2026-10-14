import { zipSync, strToU8, strFromU8 } from 'fflate';
import { Workbook } from '../dist/index.mjs';

function xlsxWithStyles(sheet, styles) {
  const ct = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`;
  const rels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const wb = `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const wbrels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  return zipSync({ '[Content_Types].xml': strToU8(ct), '_rels/.rels': strToU8(rels), 'xl/workbook.xml': strToU8(wb), 'xl/_rels/workbook.xml.rels': strToU8(wbrels), 'xl/worksheets/sheet1.xml': strToU8(sheet), 'xl/styles.xml': strToU8(styles) });
}

console.log('=== 20. custom accounting format [Red] misdetected as date ===');
{
  const styles = `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.00;[Red]-0.00"/></numFmts><cellXfs count="1"><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;
  const sheet = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="0"><v>3.14</v></c></row></sheetData></worksheet>`;
  const v = Workbook.read(xlsxWithStyles(sheet, styles)).getWorksheet('S').getCell('A1');
  console.log('A1 value =', v.value, '| type =', v.type, v.value instanceof Date ? '*** MISREAD AS DATE ***' : '(number)');
}

console.log('\n=== 21. locale-prefixed date format not detected (false negative) ===');
{
  const styles = `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="165" formatCode="[$-409]d-mmm-yy;@"/></numFmts><cellXfs count="1"><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;
  const sheet = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="0"><v>45000</v></c></row></sheetData></worksheet>`;
  const v = Workbook.read(xlsxWithStyles(sheet, styles)).getWorksheet('S').getCell('A1');
  console.log('A1 value =', v.value, '| type =', v.type, '(expected a Date)');
}

console.log('\n=== 22. CDATA in shared string gets entity-decoded ===');
{
  const ct = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`;
  const rels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const wb = `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const wbrels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;
  const ss = `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t><![CDATA[a &amp; b]]></t></si></sst>`;
  const sheet = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`;
  const x = zipSync({ '[Content_Types].xml': strToU8(ct), '_rels/.rels': strToU8(rels), 'xl/workbook.xml': strToU8(wb), 'xl/_rels/workbook.xml.rels': strToU8(wbrels), 'xl/worksheets/sheet1.xml': strToU8(sheet), 'xl/sharedStrings.xml': strToU8(ss) });
  console.log('A1 =', JSON.stringify(Workbook.read(x).getWorksheet('S').getCell('A1').value), '(XML CDATA literal should be "a &amp; b")');
}
