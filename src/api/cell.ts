import { typeOf, type CellValue, type ValueTypeName } from '../model/value';
import type { AlignmentProps, BorderProps, FillProps, FontProps } from '../model/style';
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

  get numFmt(): string {
    return this.worksheet.getCellNumFmt(this.ref);
  }

  set numFmt(value: string) {
    this.worksheet.setCellStyle(this.ref, { numFmt: value });
  }

  get font(): FontProps {
    return this.worksheet.getCellFont(this.ref);
  }

  set font(value: FontProps) {
    this.worksheet.setCellStyle(this.ref, { font: value });
  }

  get fill(): FillProps | undefined {
    return this.worksheet.getCellFill(this.ref);
  }

  set fill(value: FillProps) {
    this.worksheet.setCellStyle(this.ref, { fill: value });
  }

  get border(): BorderProps {
    return this.worksheet.getCellBorder(this.ref);
  }

  set border(value: BorderProps) {
    this.worksheet.setCellStyle(this.ref, { border: value });
  }

  get alignment(): AlignmentProps {
    return this.worksheet.getCellAlignment(this.ref);
  }

  set alignment(value: AlignmentProps) {
    this.worksheet.setCellStyle(this.ref, { alignment: value });
  }

  toString(): string {
    const value = this.value;
    return value === null || value === undefined ? '' : String(value);
  }
}
