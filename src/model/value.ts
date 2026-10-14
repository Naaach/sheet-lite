/**
 * Cell value model, aligned with ExcelJS's `CellValue` union so migration is
 * straightforward.
 */
import type { FontProps } from './style';

export type FormulaValue = { formula: string; result?: CellValue };
export type ErrorValue = { error: string };
/** One run of a rich-text string. `font` mirrors `FontProps` (name → `rFont`). */
export type RichTextRun = { text: string; font?: FontProps };
export type RichTextValue = { richText: RichTextRun[] };
export type CellValue =
  | null
  | number
  | string
  | boolean
  | Date
  | ErrorValue
  | FormulaValue
  | RichTextValue;

export const ValueType = {
  Null: 0,
  Number: 1,
  String: 2,
  Date: 3,
  Boolean: 4,
  Formula: 5,
  Error: 6,
} as const;

export type ValueTypeName = keyof typeof ValueType;

export function typeOf(value: CellValue): ValueTypeName {
  if (value === null || value === undefined) return 'Null';
  if (typeof value === 'number') return 'Number';
  if (typeof value === 'string') return 'String';
  if (typeof value === 'boolean') return 'Boolean';
  if (value instanceof Date) return 'Date';
  if (typeof value === 'object' && 'richText' in value) return 'String';
  if (typeof value === 'object' && 'formula' in value) return 'Formula';
  if (typeof value === 'object' && 'error' in value) return 'Error';
  return 'Null';
}
