import { numberToColumn } from '../reader/sheet-parser';
import type { Worksheet } from './worksheet';

/**
 * A worksheet column (ExcelJS-style). `width` and `hidden` are get/set and
 * persist through the `<cols>` element without touching any other column.
 */
export class Column {
  constructor(
    private readonly worksheet: Worksheet,
    /** 1-based column number (1 = A). */
    readonly number: number,
  ) {}

  get letter(): string {
    return numberToColumn(this.number);
  }

  get width(): number | undefined {
    return this.worksheet.getColumnWidth(this.number);
  }

  set width(value: number | undefined) {
    this.worksheet.setColumnWidth(this.number, value);
  }

  get hidden(): boolean {
    return this.worksheet.getColumnHidden(this.number);
  }

  set hidden(value: boolean) {
    this.worksheet.setColumnHidden(this.number, value);
  }
}