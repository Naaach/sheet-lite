import { strFromU8 } from 'fflate';
import type { PartStore } from '../opc/part-store';
import { resolvePartPath } from '../opc/relationships';
import { SheetLiteError } from '../zip/errors';
import { parseXml, child, childrenOf, descendants, textOf } from '../xml/xml';
import { isBuiltinDateFormat, isDateFormatCode } from '../model/dates';

export interface SheetRef {
  name: string;
  sheetId: string;
  rId: string;
  /** Absolute part name, e.g. "/xl/worksheets/sheet1.xml". */
  partName: string;
}

export interface PackageMeta {
  date1904: boolean;
  sheets: SheetRef[];
  sharedStrings: string[];
  /** numFmtId per index into `cellXfs`. */
  cellXfs: number[];
  numFmts: Map<number, string>;
  /** Indexes into `cellXfs` whose format represents a date/time. */
  dateStyleIndexes: Set<number>;
}

function decode(data: Uint8Array): string {
  return strFromU8(data);
}

function rootElement(xml: string, name: string) {
  const doc = parseXml(xml);
  return doc.children.find((n) => n.name === name) ?? doc;
}

function readRels(store: PartStore, relsPath: string, basePart: string): Map<string, string> {
  const map = new Map<string, string>();
  const data = store.get(relsPath);
  if (!data) return map;
  for (const rel of descendants(parseXml(decode(data)), 'Relationship')) {
    const id = rel.attrs.Id;
    const target = rel.attrs.Target;
    if (!id || !target) continue;
    map.set(id, resolvePartPath(basePart, target));
  }
  return map;
}

export function readPackageMeta(store: PartStore): PackageMeta {
  const workbookData = store.get('xl/workbook.xml');
  if (!workbookData) {
    throw new SheetLiteError('Not an XLSX package: missing xl/workbook.xml', 'NOT_XLSX');
  }

  const workbook = rootElement(decode(workbookData), 'workbook');
  const workbookPr = child(workbook, 'workbookPr');
  const date1904 = workbookPr?.attrs.date1904 === '1' || workbookPr?.attrs.date1904 === 'true';

  const rels = readRels(store, 'xl/_rels/workbook.xml.rels', '/xl/workbook.xml');
  const sheets: SheetRef[] = [];
  const sheetsNode = child(workbook, 'sheets');
  for (const sheet of childrenOf(sheetsNode ?? workbook, 'sheet')) {
    const rId = sheet.attrs['r:id'] ?? '';
    const partName = rels.get(rId);
    if (!partName) continue;
    sheets.push({
      name: sheet.attrs.name ?? '',
      sheetId: sheet.attrs.sheetId ?? '',
      rId,
      partName,
    });
  }

  const sharedStrings: string[] = [];
  const ssData = store.get('xl/sharedStrings.xml');
  if (ssData) {
    const sst = rootElement(decode(ssData), 'sst');
    for (const si of childrenOf(sst, 'si')) sharedStrings.push(textOf(si));
  }

  const cellXfs: number[] = [];
  const numFmts = new Map<number, string>();
  const stylesData = store.get('xl/styles.xml');
  if (stylesData) {
    const styleSheet = rootElement(decode(stylesData), 'styleSheet');
    const numFmtsNode = child(styleSheet, 'numFmts');
    if (numFmtsNode) {
      for (const nf of childrenOf(numFmtsNode, 'numFmt')) {
        numFmts.set(Number(nf.attrs.numFmtId), nf.attrs.formatCode ?? '');
      }
    }
    const xfsNode = child(styleSheet, 'cellXfs');
    if (xfsNode) {
      for (const xf of childrenOf(xfsNode, 'xf')) {
        cellXfs.push(Number(xf.attrs.numFmtId ?? '0'));
      }
    }
  }

  const dateStyleIndexes = new Set<number>();
  cellXfs.forEach((numFmtId, index) => {
    if (isBuiltinDateFormat(numFmtId)) {
      dateStyleIndexes.add(index);
      return;
    }
    const code = numFmts.get(numFmtId);
    if (code && isDateFormatCode(code)) dateStyleIndexes.add(index);
  });

  return { date1904, sheets, sharedStrings, cellXfs, numFmts, dateStyleIndexes };
}
