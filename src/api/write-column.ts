/**
 * A column of a worksheet being created.
 *
 * Width and visibility live in the `<cols>` container, which precedes
 * `<sheetData>` in the worksheet schema — so they must be set BEFORE the first
 * row is written (`PROLOGUE_LOCKED` afterwards). This mirrors the read-side
 * `Column` API (`ws.getColumn(n).width`).
 */
import { columnToNumber, numberToColumn } from '../model/refs';
import { SheetLiteError } from '../zip/errors';
import type { WriteWorksheet } from './write-worksheet';

export class WriteColumn {
  /** 1-based column index. */
  readonly index: number;

  constructor(
    private readonly worksheet: WriteWorksheet,
    col: number | string,
  ) {
    const n = typeof col === 'number' ? Math.floor(col) : columnToNumber(col);
    if (!Number.isFinite(n) || n < 1 || n > 16_384) {
      throw new SheetLiteError(
        `Invalid column ${String(col)}: must be a number 1-16384 or a letter (A-XFD)`,
        'BAD_COLUMN',
      );
    }
    this.index = n;
  }

  /** Column letters, e.g. "A", "Z", "AA" … "XFD". */
  get letter(): string {
    return numberToColumn(this.index);
  }

  get width(): number | undefined {
    return this.worksheet.getColumnDef(this.index).width;
  }

  set width(value: number | undefined) {
    this.worksheet.setColumnDef(this.index, { width: value });
  }

  get hidden(): boolean {
    return this.worksheet.getColumnDef(this.index).hidden === true;
  }

  set hidden(value: boolean) {
    this.worksheet.setColumnDef(this.index, { hidden: value ? true : undefined });
  }
}
