import { dateToSerial } from '../model/dates';
import { escapeXml } from '../xml/xml';
import type { CellValue } from '../model/value';
import { refToCoords, type ParsedSheet } from '../reader/sheet-parser';

/** Serialize a single cell into OOXML `<c>` markup. `''` removes the cell. */
export function generateCellXml(
  ref: string,
  value: CellValue,
  styleIndex: number | undefined,
  date1904: boolean,
): string {
  if (value === null || value === undefined) return '';
  const sAttr = styleIndex ? ` s="${styleIndex}"` : '';

  if (typeof value === 'number') {
    return `<c r="${ref}"${sAttr}><v>${value}</v></c>`;
  }
  if (typeof value === 'boolean') {
    return `<c r="${ref}" t="b"${sAttr}><v>${value ? 1 : 0}</v></c>`;
  }
  if (value instanceof Date) {
    return `<c r="${ref}"${sAttr}><v>${dateToSerial(value, date1904)}</v></c>`;
  }
  if (typeof value === 'object' && 'error' in value) {
    return `<c r="${ref}" t="e"${sAttr}><v>${escapeXml(value.error)}</v></c>`;
  }
  if (typeof value === 'object' && 'formula' in value) {
    const { result } = value;
    let resultText = '';
    if (result instanceof Date) {
      resultText = String(dateToSerial(result, date1904));
    } else if (
      typeof result === 'number' ||
      typeof result === 'string' ||
      typeof result === 'boolean'
    ) {
      resultText = String(result);
    }
    return `<c r="${ref}"${sAttr}><f>${escapeXml(value.formula)}</f><v>${escapeXml(resultText)}</v></c>`;
  }

  const str = String(value);
  const preserve = /^\s|\s$|[\n\t]/.test(str) ? ' xml:space="preserve"' : '';
  return `<c r="${ref}" t="inlineStr"${sAttr}><is><t${preserve}>${escapeXml(str)}</t></is></c>`;
}

interface Op {
  pos: number;
  end: number;
  text: string;
}

/**
 * Patch a worksheet: only the changed cells are rewritten; every untouched byte
 * of the original XML (styles, validations, drawings, conditional formatting, …)
 * stays exactly as it was.
 */
export function serializeSheet(
  sheet: ParsedSheet,
  changes: Map<string, CellValue>,
  date1904: boolean,
): string {
  const ops: Op[] = [];
  const insertsByPos = new Map<number, { col: number; xml: string }[]>();
  const newRows = new Map<number, { col: number; xml: string }[]>();

  for (const [ref, value] of changes) {
    const existing = sheet.cells.get(ref);
    if (existing) {
      ops.push({
        pos: existing.start,
        end: existing.end,
        text: generateCellXml(ref, value, existing.styleIndex, date1904),
      });
      continue;
    }

    const { row, col } = refToCoords(ref);
    const xml = generateCellXml(ref, value, undefined, date1904);
    const rowRec = sheet.rows.get(row);
    if (rowRec) {
      const list = insertsByPos.get(rowRec.closeStart) ?? [];
      list.push({ col, xml });
      insertsByPos.set(rowRec.closeStart, list);
    } else {
      const list = newRows.get(row) ?? [];
      list.push({ col, xml });
      newRows.set(row, list);
    }
  }

  for (const [pos, list] of insertsByPos) {
    list.sort((a, b) => a.col - b.col);
    ops.push({ pos, end: pos, text: list.map((entry) => entry.xml).join('') });
  }

  const newRowXml: string[] = [];
  for (const rowNum of [...newRows.keys()].sort((a, b) => a - b)) {
    const list = newRows.get(rowNum)!;
    list.sort((a, b) => a.col - b.col);
    newRowXml.push(`<row r="${rowNum}">${list.map((entry) => entry.xml).join('')}</row>`);
  }
  if (newRowXml.length > 0) {
    ops.push({ pos: sheet.sheetDataClose, end: sheet.sheetDataClose, text: newRowXml.join('') });
  }

  ops.sort((a, b) => b.pos - a.pos);
  let out = sheet.xml;
  for (const op of ops) {
    out = out.slice(0, op.pos) + op.text + out.slice(op.end);
  }
  return out;
}
