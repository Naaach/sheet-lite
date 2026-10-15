/** Basic style properties, aligned with ExcelJS's shapes where practical. */

export interface FontProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  size?: number;
  name?: string;
  /** ARGB like "FFFF0000". */
  color?: string;
}

export interface FillProps {
  type?: 'pattern';
  pattern?: string;
  /** ARGB like "FFFFFF00". */
  fgColor?: string;
  bgColor?: string;
}

/** Line style of one border edge, matching the OOXML `borderStyle` values. */
export type BorderStyleName =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'slantDashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot';

/** One border edge, ExcelJS-shaped (`color.argb` like "FF0000FF"). */
export interface BorderSideProps {
  style?: BorderStyleName;
  color?: { argb?: string };
}

export interface BorderProps {
  top?: BorderSideProps;
  bottom?: BorderSideProps;
  left?: BorderSideProps;
  right?: BorderSideProps;
  diagonal?: BorderSideProps;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

/** Cell alignment; lives inside the `xf` as an `<alignment/>` child. */
export interface AlignmentProps {
  horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
  /** OOXML textRotation: 0-90, 91-180 (counter-clockwise) or 255 (vertical). */
  textRotation?: number;
  /** OOXML readingOrder: 0 = context, 1 = left-to-right, 2 = right-to-left. */
  readingOrder?: 0 | 1 | 2;
}

export interface CellStyleOverrides {
  numFmt?: string;
  font?: FontProps;
  fill?: FillProps;
  border?: BorderProps;
  alignment?: AlignmentProps;
}

/**
 * Merge a style patch over an existing one, property by property (a `font`
 * patch only replaces the font fields it names, it never drops the others).
 * Shared by the incremental writer, where a row-level patch and a cell-level
 * patch combine.
 */
export function mergeStylePatch(
  base: CellStyleOverrides,
  patch: CellStyleOverrides,
): CellStyleOverrides {
  const out: CellStyleOverrides = { ...base };
  if (patch.numFmt !== undefined) out.numFmt = patch.numFmt;
  if (patch.font) out.font = { ...out.font, ...patch.font };
  if (patch.fill) out.fill = { ...out.fill, ...patch.fill };
  if (patch.border) out.border = { ...out.border, ...patch.border };
  if (patch.alignment) out.alignment = { ...out.alignment, ...patch.alignment };
  return out;
}
