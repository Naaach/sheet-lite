import { typeOf, type CellValue, type ValueTypeName } from '../model/value';
import { refToCoords } from '../reader/sheet-parser';
import type { Worksheet } from './worksheet';

export class Cell {
  constructor(
    private readonly worksheet: Worksheet,
    readonly ref: string,
  ) {}

  get address(): string {
    return this.ref;
  }

  get row(): number {
    return refToCoords(this.ref).row;
  }

  get col(): number {
    return refToCoords(this.ref).col;
  }

  get value(): CellValue {
    return this.worksheet.getCellValue(this.ref);
  }

  set value(value: CellValue) {
    this.worksheet.setCellValue(this.ref, value);
  }

  get type(): ValueTypeName {
    return typeOf(this.value);
  }

  get model(): { address: string; type: ValueTypeName; value: CellValue } {
    return { address: this.ref, type: this.type, value: this.value };
  }

  toString(): string {
    const value = this.value;
    return value === null || value === undefined ? '' : String(value);
  }
}
