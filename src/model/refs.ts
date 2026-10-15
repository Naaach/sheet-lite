/**
 * A1-style cell references and column letters.
 *
 * Pure leaf helpers, shared by the sheet reader, the surgical writer and the
 * incremental writer — no DOM, no ZIP, no OPC.
 */
import { SheetLiteError } from '../zip/errors';

/** Excel's hard limits, mirroring `MAX_ROWS`/`MAX_COLUMNS` in the sheet writer. */
const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

/** `"A"`, `"Z"`, `"aa"` … -> 1, 26, 27 … */
export function columnToNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** 1, 26, 27 … -> `"A"`, `"Z"`, `"AA"` … */
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

/** `"B7"` -> `{ row: 7, col: 2 }`; anything unparsable maps to `{ row: 0, col: 0 }`. */
export function refToCoords(ref: string): { row: number; col: number } {
  const match = /^([A-Za-z]+)(\d+)$/.exec(ref);
  if (!match) return { row: 0, col: 0 };
  return { row: Number(match[2]), col: columnToNumber(match[1]!) };
}

/** `{ row: 7, col: 2 }` -> `"B7"`. */
export function coordsToRef(row: number, col: number): string {
  return `${numberToColumn(col)}${row}`;
}

const badMergeRange = (ref: string): never => {
  throw new SheetLiteError(
    `Invalid merge range ${JSON.stringify(ref)}: must be a cell range like A1:C1 (rows 1-${MAX_ROW}, columns A-XFD)`,
    'BAD_MERGE_REF',
  );
};

/**
 * Validate a merged range and return it normalized (upper-case, one or two
 * endpoints). `<mergeCell ref="…">` is written verbatim, so anything that is
 * not a real range — `NOT-A-RANGE!!`, a row past XFD — is a file Excel offers
 * to repair. Shared by both writers: the eager one when the sheet is
 * serialized, the incremental one when `mergeCells()` is called.
 *
 * @throws {SheetLiteError} `BAD_MERGE_REF`.
 */
export function assertMergeRange(ref: string): string {
  const upper = String(ref).toUpperCase();
  const parts = upper.split(':');
  if (parts.length > 2) badMergeRange(ref);
  for (const part of parts) {
    const match = /^([A-Z]{1,3})(\d{1,7})$/.exec(part);
    const row = match ? Number(match[2]) : 0;
    if (!match || row < 1 || row > MAX_ROW || columnToNumber(match[1]!) > MAX_COL) badMergeRange(ref);
  }
  return upper;
}