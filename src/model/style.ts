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

export interface CellStyleOverrides {
  numFmt?: string;
  font?: FontProps;
  fill?: FillProps;
}
