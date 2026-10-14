import { zipSync, strToU8, strFromU8 } from 'fflate';
import { Workbook } from '../dist/index.mjs';

function minimalXlsx() {
  const ct = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`;
  const rels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  const wb = `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const wbrels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;
  const sheet = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`;
  return zipSync({ '[Content_Types].xml': strToU8(ct), '_rels/.rels': strToU8(rels), 'xl/workbook.xml': strToU8(wb), 'xl/_rels/workbook.xml.rels': strToU8(wbrels), 'xl/worksheets/sheet1.xml': strToU8(sheet) });
}

console.log('=== 17. low-level edit of [Content_Types].xml only ===');
{
  const x = minimalXlsx();
  const wb = Workbook.read(x);
  const original = strFromU8(wb.parts().get('[Content_Types].xml'));
  const edited = original.replace('application/xml', 'application/edited+xml');
  wb.parts().set('[Content_Types].xml', strToU8(edited));
  const out = wb.toBuffer();
  const after = strFromU8(Workbook.read(out).parts().get('[Content_Types].xml'));
  console.log('edit persisted?', after.includes('application/edited+xml'));
  console.log('output identical to original bytes?', Buffer.from(out).equals(Buffer.from(x)));
}

console.log('\n=== 18. low-level edit of a non-Content_Types part ===');
{
  const x = minimalXlsx();
  const wb = Workbook.read(x);
  wb.parts().set('docProps/custom.xml', strToU8('<x/>'));
  const out = wb.toBuffer();
  console.log('new part persisted?', !!Workbook.read(out).parts().get('docProps/custom.xml'));
}

console.log('\n=== 19. sheet dirty + content-types dirty together ===');
{
  const x = minimalXlsx();
  const wb = Workbook.read(x);
  wb.parts().set('[Content_Types].xml', strToU8(strFromU8(wb.parts().get('[Content_Types].xml')).replace('application/xml', 'application/edited+xml')));
  wb.getWorksheet('S').getCell('A1').value = 99;
  const out = wb.toBuffer();
  const after = strFromU8(Workbook.read(out).parts().get('[Content_Types].xml'));
  console.log('CT edit persisted?', after.includes('application/edited+xml'));
  console.log('sheet edit persisted?', Workbook.read(out).getWorksheet('S').getCell('A1').value === 99);
}
