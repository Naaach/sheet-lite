/**
 * Style-value validation, shared by BOTH writers.
 *
 * `StyleRegistry` (eager, `writer/style-registry.ts`) and `StylesBuilder`
 * (incremental, `writer/style-builder.ts`) serialize the same
 * `CellStyleOverrides` shapes into the same `styles.xml` fragments, so they
 * must accept and reject exactly the same values. Every attribute checked here
 * is an `xsd:int`, an enumeration or a hex string in the SpreadsheetML schema:
 * a value outside its domain is written verbatim into `<sz val>`, `<alignment
 * indent>` or `<color rgb>` and is a file Excel offers to repair.
 *
 * Free-form strings (a font name, a fill pattern) are the writers' other half:
 * they are escaped, not validated.
 *
 * @throws {SheetLiteError} `BAD_STYLE` for any value the schema forbids.
 */
import { SheetLiteError } from '../zip/errors';
import type { BorderProps, CellStyleOverrides, FontProps } from '../model/style';

/** The 13 `ST_BorderStyle` values, in schema order (space-separated to match). */
const BORDER_STYLES =
  'thin medium thick dashed dotted double hair mediumDashed dashDot slantDashDot mediumDashDot dashDotDot mediumDashDotDot';
const BORDER_STYLES_PADDED = ` ${BORDER_STYLES} `;

/** `ST_UnsignedIntHex`: exactly 8 hex digits, i.e. an ARGB colour. */
const ARGB = /^[0-9a-f]{8}$/i;

/** `xsd:int` in [min,max] — which also rejects NaN, Infinity and fractions. */
const int = (value: number, min: number, max: number): boolean =>
  Number.isInteger(value) && value >= min && value <= max;

function bad(what: string, value: string, rule: string): never {
  throw new SheetLiteError(`Invalid ${what} ${value}: ${rule}`, 'BAD_STYLE');
}

function color(what: string, value: unknown): void {
  if (value !== undefined && !(typeof value === 'string' && ARGB.test(value))) {
    bad(what, JSON.stringify(value) ?? String(value), 'must be an ARGB color of 8 hex digits, e.g. "FFFF0000"');
  }
}

/**
 * A font's numeric and colour attributes (`<sz val>`, `<color rgb>`), for a
 * cell font as well as for a rich-text run's `<rPr>`.
 */
export function assertFont(font: FontProps): void {
  if (font.size !== undefined && !int(font.size, 1, 409)) {
    bad('font size', String(font.size), 'must be an integer 1-409');
  }
  color('font color', font.color);
}

/** One border edge: a `ST_BorderStyle` name and an optional ARGB colour. */
function assertBorder(border: BorderProps): void {
  for (const name of ['top', 'bottom', 'left', 'right', 'diagonal'] as const) {
    const edge = border[name];
    if (!edge) continue;
    if (edge.style !== undefined && !BORDER_STYLES_PADDED.includes(` ${edge.style} `)) {
      bad(`border ${name} style`, JSON.stringify(edge.style), `must be one of ${BORDER_STYLES}`);
    }
    color(`border ${name} color`, edge.color?.argb);
  }
}

/**
 * Validate one cell's style patch — exactly the model the writers receive from
 * `Cell.font/fill/border/alignment` and from `WriteRow`/`WriteCell`.
 *
 * @throws {SheetLiteError} `BAD_STYLE` for any value the schema forbids.
 */
export function assertStylePatch(patch: CellStyleOverrides): void {
  if (patch.font) assertFont(patch.font);
  if (patch.fill) {
    color('fill fgColor', patch.fill.fgColor);
    color('fill bgColor', patch.fill.bgColor);
  }
  if (patch.border) assertBorder(patch.border);
  const alignment = patch.alignment;
  if (!alignment) return;
  if (alignment.textRotation !== undefined && !(int(alignment.textRotation, 0, 180) || alignment.textRotation === 255)) {
    bad('alignment textRotation', String(alignment.textRotation), 'must be an integer 0-180, or 255 for vertical text');
  }
  if (alignment.indent !== undefined && !int(alignment.indent, 0, 250)) {
    bad('alignment indent', String(alignment.indent), 'must be an integer 0-250');
  }
  if (alignment.readingOrder !== undefined && !int(alignment.readingOrder, 0, 2)) {
    bad('alignment readingOrder', String(alignment.readingOrder), 'must be 0 (context), 1 (left-to-right) or 2 (right-to-left)');
  }
}