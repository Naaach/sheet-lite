import { parseXml, child, childrenOf, textOf, type XmlNode } from '../xml/xml';
import { serialToDate } from '../model/dates';
import { typeOf, type CellValue, type ValueTypeName } from '../model/value';
import type { PackageMeta } from './package-reader';

export interface CellRecord {
  ref: string;
  row: number;
  col: number;
  /** Byte range of the `<c>` element in the original sheet XML. */
  start: number;
  end: number;
  styleIndex: number;
  value: CellValue;
  type: ValueTypeName;
}

export interface RowRecord {
  r: number;
  start: number;
  closeStart: number;
  end: number;
}

export interface ParsedSheet {
  xml: string;
  cells: Map<string, CellRecord>;
  rows: Map<number, RowRecord>;
  /** Index of the `</sheetData>` close tag in the original XML. */
  sheetDataClose: number;
  maxRow: number;
  maxCol: number;
}

export function columnToNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numberToColumn(n: number): string {
  let s = '';
  let value = n;
  while (value > 0) {
    const rem = (value - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    value = Math.floor((value - 1) / 26);
  }
  return s;
}

export function refToCoords(ref: string): { row: number; col: number } {
  const match = /^([A-Za-z]+)(\d+)$/.exec(ref);
  if (!match) return { row: 0, col: 0 };
  return { row: Number(match[2]), col: columnToNumber(match[1]!) };
}

export function coordsToRef(row: number, col: number): string {
  return `${numberToColumn(col)}${row}`;
}

function computeValue(c: XmlNode, meta: PackageMeta): CellValue {
  const t = c.attrs.t;
  const styleIndex = c.attrs.s ? Number(c.attrs.s) : 0;
  const vNode = child(c, 'v');
  const rawValue = vNode ? textOf(vNode) : undefined;
  const fNode = child(c, 'f');

  let value: CellValue = null;
  if (t === 's') {
    value = meta.sharedStrings[Number(rawValue)] ?? '';
  } else if (t === 'inlineStr') {
    const isNode = child(c, 'is');
    value = isNode ? textOf(isNode) : '';
  } else if (t === 'str') {
    value = rawValue ?? '';
  } else if (t === 'b') {
    value = rawValue === '1';
  } else if (t === 'e') {
    value = { error: rawValue ?? '' };
  } else if (rawValue !== undefined) {
    const numeric = Number(rawValue);
    value = meta.dateStyleIndexes.has(styleIndex)
      ? serialToDate(numeric, meta.date1904)
      : numeric;
  }

  if (fNode) {
    return { formula: textOf(fNode), result: value ?? undefined };
  }
  return value;
}

export function parseSheet(xml: string, meta: PackageMeta): ParsedSheet {
  const document = parseXml(xml);
  const worksheet = document.children.find((n) => n.name === 'worksheet') ?? document;
  const sheetData = child(worksheet, 'sheetData');

  const cells = new Map<string, CellRecord>();
  const rows = new Map<number, RowRecord>();
  let maxRow = 0;
  let maxCol = 0;

  if (sheetData) {
    let implicitRow = 0;
    for (const rowNode of childrenOf(sheetData, 'row')) {
      implicitRow += 1;
      const r = rowNode.attrs.r ? Number(rowNode.attrs.r) : implicitRow;
      rows.set(r, {
        r,
        start: rowNode.start,
        closeStart: rowNode.closeStart,
        end: rowNode.end,
      });
      if (r > maxRow) maxRow = r;

      for (const c of childrenOf(rowNode, 'c')) {
        const ref = c.attrs.r;
        if (!ref) continue;
        const { row, col } = refToCoords(ref);
        const value = computeValue(c, meta);
        cells.set(ref, {
          ref,
          row,
          col,
          start: c.start,
          end: c.end,
          styleIndex: c.attrs.s ? Number(c.attrs.s) : 0,
          value,
          type: typeOf(value),
        });
        if (col > maxCol) maxCol = col;
        if (row > maxRow) maxRow = row;
      }
    }
  }

  const closeIdx = xml.lastIndexOf('</sheetData>');
  return {
    xml,
    cells,
    rows,
    sheetDataClose: closeIdx === -1 ? xml.lastIndexOf('</worksheet>') : closeIdx,
    maxRow,
    maxCol,
  };
}
