import { dateToSerial } from '../model/dates';
import { escapeXml } from '../xml/xml';
import type { CellValue } from '../model/value';
import type { CellStyleOverrides } from '../model/style';
import { refToCoords, type ParsedSheet } from '../reader/sheet-parser';
import type { StyleRegistry } from './style-registry';

export interface CellEdit {
  hasValue: boolean;
  value: CellValue;
  style?: CellStyleOverrides;
}

/** Serialize a single cell into OOXML `<c>` markup. `''` removes the cell. */
export function generateCellXml(
  ref: string,
  value: CellValue,
  styleIndex: number,
  date1904: boolean,
): string {
  if (value === null || value === undefined) return '';
  const sAttr = styleIndex ? ` s="${styleIndex}"` : '';

  if (typeof value === 'number') return `<c r="${ref}"${sAttr}><v>${value}</v></c>`;
  if (typeof value === 'boolean') return `<c r="${ref}" t="b"${sAttr}><v>${value ? 1 : 0}</v></c>`;
  if (value instanceof Date) return `<c r="${ref}"${sAttr}><v>${dateToSerial(value, date1904)}</v></c>`;
  if (typeof value === 'object' && 'error' in value) {
    return `<c r="${ref}" t="e"${sAttr}><v>${escapeXml(value.error)}</v></c>`;
  }
  if (typeof value === 'object' && 'formula' in value) {
    const { result } = value;
    let resultText = '';
    if (result instanceof Date) resultText = String(dateToSerial(result, date1904));
    else if (typeof result === 'number' || typeof result === 'string' || typeof result === 'boolean') {
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
 * Patch a worksheet: only the changed cells/merges are rewritten; every
 * untouched byte of the original XML stays exactly as it was.
 */
export function serializeSheet(
  sheet: ParsedSheet,
  changes: Map<string, CellEdit>,
  date1904: boolean,
  registry: StyleRegistry | null,
  newMerges: string[],
  removedMerges: string[],
): string {
  const ops: Op[] = [];
  const insertsByPos = new Map<number, { col: number; xml: string }[]>();
  const newRows = new Map<number, { col: number; xml: string }[]>();
  const selfClosingRows = new Map<number, { col: number; xml: string }[]>();

  for (const [ref, edit] of changes) {
    const existing = sheet.cells.get(ref);
    const baseIndex = existing ? existing.styleIndex : 0;
    const styleIndex =
      edit.style && registry ? registry.resolveIndex(baseIndex, edit.style) : baseIndex;
    const value = edit.hasValue ? edit.value : existing ? existing.value : null;
    let xml: string;
    if ((value === null || value === undefined) && edit.style && styleIndex) {
      // An empty cell that only carries styling still needs to exist.
      xml = `<c r="${ref}" s="${styleIndex}"/>`;
    } else {
      xml = generateCellXml(ref, value, styleIndex, date1904);
    }

    if (existing) {
      ops.push({ pos: existing.start, end: existing.end, text: xml });
      continue;
    }
    const { row, col } = refToCoords(ref);
    const rowRec = sheet.rows.get(row);
    if (rowRec && !rowRec.selfClosing) {
      const list = insertsByPos.get(rowRec.closeStart) ?? [];
      list.push({ col, xml });
      insertsByPos.set(rowRec.closeStart, list);
    } else if (rowRec && rowRec.selfClosing) {
      const list = selfClosingRows.get(row) ?? [];
      list.push({ col, xml });
      selfClosingRows.set(row, list);
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

  for (const [rowNum, list] of selfClosingRows) {
    list.sort((a, b) => a.col - b.col);
    const rowRec = sheet.rows.get(rowNum)!;
    const openTag = sheet.xml.slice(rowRec.start, rowRec.openEnd).replace(/\/>\s*$/, '>');
    ops.push({
      pos: rowRec.start,
      end: rowRec.end,
      text: openTag + list.map((entry) => entry.xml).join('') + '</row>',
    });
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

  applyMergeEdits(sheet, newMerges, removedMerges, ops);

  ops.sort((a, b) => b.pos - a.pos);
  let out = sheet.xml;
  for (const op of ops) out = out.slice(0, op.pos) + op.text + out.slice(op.end);
  return out;
}

function applyMergeEdits(sheet: ParsedSheet, newMerges: string[], removedMerges: string[], ops: Op[]): void {
  const container = sheet.mergeCells;
  let present = sheet.mergeChildren;

  const removedSet = new Set(removedMerges);
  if (removedMerges.length > 0) {
    for (const child of present) {
      if (removedSet.has(child.ref)) ops.push({ pos: child.start, end: child.end, text: '' });
    }
    present = present.filter((child) => !removedSet.has(child.ref));
  }

  const hasAdd = newMerges.length > 0;
  const additions = newMerges.map((ref) => `<mergeCell ref="${ref}"/>`).join('');

  if (container) {
    if (!hasAdd && removedMerges.length === 0) return;
    const count = present.length + newMerges.length;
    const open = sheet.xml.slice(container.start, container.openEnd);
    const updated = /\bcount="\d+"/.test(open)
      ? open.replace(/\bcount="\d+"/, `count="${count}"`)
      : open.replace(/>$/, ` count="${count}">`);
    ops.push({ pos: container.start, end: container.openEnd, text: updated });
    // Additions go at the end of the container; removals are separate range ops.
    if (hasAdd) ops.push({ pos: container.closeStart, end: container.closeStart, text: additions });
  } else if (hasAdd) {
    // mergeCells is a sibling of sheetData, so it goes AFTER </sheetData>.
    ops.push({
      pos: sheet.sheetDataEnd,
      end: sheet.sheetDataEnd,
      text: `<mergeCells count="${newMerges.length}">${additions}</mergeCells>`,
    });
  }
}
