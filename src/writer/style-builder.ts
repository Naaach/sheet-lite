/**
 * Incremental `styles.xml` builder for the streaming writer.
 *
 * The eager `StyleRegistry` (`writer/style-registry.ts`) EXTENDS a styles.xml
 * it found in a file, patching only what changed; this one builds the sheet from
 * nothing and appends to it as cells are written. Both emit the same containers,
 * the same child order and the same attribute order for the same style, and
 * both dedup: a style used on a million cells costs one entry.
 *
 * It deliberately does NOT import the eager registry (which would drag the XML
 * parser into the writer's kernel). What IS shared is the model: the built-in
 * number-format table (`model/dates.ts`), the `CellStyleOverrides` shapes and
 * the style-value rules (`writer/style-validate.ts`), so the two writers accept
 * and reject exactly the same styles.
 */
import { BUILTIN_NUMFMT } from '../model/dates';
import type {
  AlignmentProps,
  BorderProps,
  BorderSideProps,
  CellStyleOverrides,
  FillProps,
  FontProps,
} from '../model/style';
import { escapeXml } from '../xml/escape';
import { SHEET_NS, XML_DECL } from './opc-constants';
import { assertStylePatch } from './style-validate';

/**
 * Style sheet a workbook created from scratch starts from: the same minimal
 * baseline the eager writer creates when a file has no `styles.xml`
 * (`opc/minimal-package.ts`), so both writers emit identical containers for
 * identical styles. Everything else is appended as cells are written.
 */
const BASE_FONT = '<font><sz val="11"/><name val="Calibri"/></font>';
const BASE_FILL = '<fill><patternFill patternType="none"/></fill>';
const BASE_BORDER = '<border/>';

/** `<left/>` … `<diagonal/>` children of a `<border>`, in schema order. */
const BORDER_SIDES = ['left', 'right', 'top', 'bottom', 'diagonal'] as const;

/**
 * Builds `xl/styles.xml` as cells are written and hands out `cellXfs` indices.
 *
 * ```ts
 * const styles = new StylesBuilder();
 * cell.numFmt = '#,##0.00';          // cell gets `s="1"`
 * const xml = styles.toXml();        // numFmts + the new xf
 * ```
 */
export class StylesBuilder {
  private readonly numFmtByCode = new Map<string, number>();
  private readonly numFmts: string[] = [];
  private readonly fonts: string[] = [BASE_FONT];
  private readonly fills: string[] = [BASE_FILL];
  private readonly borders: string[] = [BASE_BORDER];
  private readonly xfs: string[] = [
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
  ];
  /** Raw entry -> index, so dedup is a hash lookup, not a scan. */
  private readonly fontIds = new Map<string, number>([[BASE_FONT, 0]]);
  private readonly fillIds = new Map<string, number>([[BASE_FILL, 0]]);
  private readonly borderIds = new Map<string, number>([[BASE_BORDER, 0]]);
  private readonly xfIds = new Map<string, number>();
  private nextNumFmtId = 164;
  private changed = false;

  /** True once at least one entry was appended. */
  get dirty(): boolean {
    return this.changed;
  }

  /**
   * Resolve a cell's style patch to a `cellXfs` index (0 = the default style).
   * Identical styles collapse onto one index, whatever the order they arrive in.
   *
   * @throws {SheetLiteError} `BAD_STYLE` for a value the schema forbids.
   */
  resolveIndex(patch: CellStyleOverrides): number {
    if (!hasOverrides(patch)) return 0;
    assertStylePatch(patch);
    const numFmtId = patch.numFmt !== undefined ? this.numFmtIdFor(patch.numFmt) : 0;
    const fontId = patch.font ? this.fontIdFor(patch.font) : 0;
    const fillId = patch.fill ? this.fillIdFor(patch.fill) : 0;
    const borderId = patch.border ? this.borderIdFor(patch.border) : 0;
    const alignmentXml = alignmentXmlFor(patch.alignment);

    // Same ids + same alignment XML == the same style, whatever order the
    // attributes were given in.
    const key = `${numFmtId} ${fontId} ${fillId} ${borderId} ${alignmentXml}`;
    const existing = this.xfIds.get(key);
    if (existing !== undefined) return existing;

    const apply =
      (patch.numFmt !== undefined ? ' applyNumberFormat="1"' : '') +
      (patch.font ? ' applyFont="1"' : '') +
      (patch.fill ? ' applyFill="1"' : '') +
      (patch.border ? ' applyBorder="1"' : '') +
      (patch.alignment ? ' applyAlignment="1"' : '');
    const attrs =
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" ` +
      `borderId="${borderId}" xfId="0"${apply}`;
    const raw = alignmentXml ? `${attrs}>${alignmentXml}</xf>` : `${attrs}/>`;
    this.xfs.push(raw);
    this.xfIds.set(key, this.xfs.length - 1);
    this.changed = true;
    return this.xfs.length - 1;
  }

  /** Serialize the whole style sheet (baseline containers + everything added). */
  toXml(): string {
    const numFmts =
      this.numFmts.length > 0
        ? `<numFmts count="${this.numFmtByCode.size}">${this.numFmts.join('')}</numFmts>`
        : '';
    return (
      XML_DECL +
      `<styleSheet ${SHEET_NS}>` +
      numFmts +
      `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join('')}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join('')}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>` +
      '</styleSheet>'
    );
  }

  /** Append `raw` to a container unless an identical entry is already there. */
  private intern(list: string[], index: Map<string, number>, raw: string): number {
    const existing = index.get(raw);
    if (existing !== undefined) return existing;
    list.push(raw);
    index.set(raw, list.length - 1);
    return list.length - 1;
  }

  private numFmtIdFor(code: string): number {
    const builtin = BUILTIN_NUMFMT[code];
    if (builtin !== undefined) return builtin;
    const existing = this.numFmtByCode.get(code);
    if (existing !== undefined) return existing;
    const id = this.nextNumFmtId++;
    this.numFmtByCode.set(code, id);
    this.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`);
    return id;
  }

  private fontIdFor(props: FontProps): number {
    const inner =
      (props.bold ? '<b/>' : '') +
      (props.italic ? '<i/>' : '') +
      (props.underline ? '<u/>' : '') +
      (props.size !== undefined ? `<sz val="${props.size}"/>` : '') +
      (props.color ? `<color rgb="${escapeXml(props.color)}"/>` : '') +
      (props.name ? `<name val="${escapeXml(props.name)}"/>` : '');
    if (inner === '') return 0;
    return this.intern(this.fonts, this.fontIds, `<font>${inner}</font>`);
  }

  private fillIdFor(props: FillProps): number {
    const inner =
      `<patternFill patternType="${escapeXml(props.pattern ?? 'solid')}">` +
      (props.fgColor ? `<fgColor rgb="${escapeXml(props.fgColor)}"/>` : '') +
      (props.bgColor ? `<bgColor rgb="${escapeXml(props.bgColor)}"/>` : '<bgColor indexed="64"/>') +
      '</patternFill>';
    return this.intern(this.fills, this.fillIds, `<fill>${inner}</fill>`);
  }

  private borderIdFor(props: BorderProps): number {
    const side = (name: string, edge: BorderSideProps | undefined): string => {
      const style = edge?.style;
      const argb = edge?.color?.argb;
      if (!style && !argb) return `<${name}/>`;
      const color = argb ? `<color rgb="${escapeXml(argb)}"/>` : '';
      if (!style) return `<${name}>${color}</${name}>`;
      return color
        ? `<${name} style="${escapeXml(style)}">${color}</${name}>`
        : `<${name} style="${escapeXml(style)}"/>`;
    };
    // Every side empty: the baseline `<border/>` already is that style.
    const bare =
      !props.diagonalUp &&
      !props.diagonalDown &&
      BORDER_SIDES.every((name) => {
        const edge = props[name];
        return !edge?.style && !edge?.color?.argb;
      });
    if (bare) return 0;
    const diagAttrs =
      (props.diagonalUp ? ' diagonalUp="1"' : '') + (props.diagonalDown ? ' diagonalDown="1"' : '');
    const raw =
      `<border${diagAttrs}>` +
      BORDER_SIDES.map((name) => side(name, props[name])).join('') +
      '</border>';
    return this.intern(this.borders, this.borderIds, raw);
  }
}

const hasOverrides = (patch: CellStyleOverrides): boolean =>
  patch.numFmt !== undefined ||
  patch.font !== undefined ||
  patch.fill !== undefined ||
  patch.border !== undefined ||
  patch.alignment !== undefined;

/** `<alignment/>` child of an `xf`, or '' when the xf carries none. */
function alignmentXmlFor(props: AlignmentProps | undefined): string {
  if (!props) return '';
  let attrs = '';
  if (props.horizontal !== undefined) attrs += ` horizontal="${escapeXml(props.horizontal)}"`;
  if (props.vertical !== undefined) attrs += ` vertical="${escapeXml(props.vertical)}"`;
  if (props.wrapText !== undefined) attrs += ` wrapText="${props.wrapText ? '1' : '0'}"`;
  if (props.shrinkToFit !== undefined) attrs += ` shrinkToFit="${props.shrinkToFit ? '1' : '0'}"`;
  if (props.indent !== undefined) attrs += ` indent="${props.indent}"`;
  if (props.textRotation !== undefined) attrs += ` textRotation="${props.textRotation}"`;
  if (props.readingOrder !== undefined) attrs += ` readingOrder="${props.readingOrder}"`;
  return attrs ? `<alignment${attrs}/>` : '';
}