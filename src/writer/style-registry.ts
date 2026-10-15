import { parseXml, child, childrenOf, decodeEntities, escapeXml, type XmlNode } from '../xml/xml';
import { SheetLiteError } from '../zip/errors';
import { BUILTIN_BY_ID, BUILTIN_NUMFMT } from '../model/dates';
import { assertStylePatch } from './style-validate';
import type {
  AlignmentProps,
  BorderProps,
  BorderSideProps,
  CellStyleOverrides,
  FillProps,
  FontProps,
} from '../model/style';

interface XfDesc {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  /** Parsed `<alignment/>` child of the xf (absent = no alignment). */
  alignmentProps?: AlignmentProps;
  /** Canonical `<alignment/>` markup ('' = no alignment): the dedup key. */
  alignmentXml: string;
  raw: string;
}

const normalize = (xml: string): string => xml.replace(/>\s+</g, '><').trim();

/** `<left/>` … `<diagonal/>` children of a `<border>`, in schema order. */
const BORDER_SIDES = ['left', 'right', 'top', 'bottom', 'diagonal'] as const;

function sideProps(node: XmlNode | undefined): BorderSideProps | undefined {
  if (!node) return undefined;
  const style = node.attrs.style;
  const rgb = child(node, 'color')?.attrs.rgb;
  if (!style && !rgb) return undefined; // `<left/>` with no style/color: nothing to expose
  const side: BorderSideProps = {};
  if (style) side.style = style as BorderSideProps['style'];
  if (rgb) side.color = { argb: rgb };
  return side;
}

function parseBorderNode(node: XmlNode): BorderProps {
  const props: BorderProps = {};
  for (const name of BORDER_SIDES) {
    const side = sideProps(child(node, name));
    if (side) props[name] = side;
  }
  if (node.attrs.diagonalUp === '1' || node.attrs.diagonalUp === 'true') props.diagonalUp = true;
  if (node.attrs.diagonalDown === '1' || node.attrs.diagonalDown === 'true') props.diagonalDown = true;
  return props;
}

function parseAlignmentNode(node: XmlNode): AlignmentProps {
  const props: AlignmentProps = {};
  const a = node.attrs;
  if (a.horizontal !== undefined) props.horizontal = a.horizontal as AlignmentProps['horizontal'];
  if (a.vertical !== undefined) props.vertical = a.vertical as AlignmentProps['vertical'];
  if (a.wrapText !== undefined) props.wrapText = a.wrapText === '1' || a.wrapText === 'true';
  if (a.shrinkToFit !== undefined) props.shrinkToFit = a.shrinkToFit === '1' || a.shrinkToFit === 'true';
  if (a.indent !== undefined) {
    const indent = Number(a.indent);
    if (Number.isFinite(indent)) props.indent = indent;
  }
  if (a.textRotation !== undefined) {
    const rotation = Number(a.textRotation);
    if (Number.isFinite(rotation)) props.textRotation = rotation;
  }
  if (a.readingOrder !== undefined) {
    const order = Number(a.readingOrder);
    if (order === 0 || order === 1 || order === 2) props.readingOrder = order;
  }
  return props;
}

function slice(xml: string, node: XmlNode): string {
  return xml.slice(node.start, node.end);
}

interface Edit {
  start: number;
  end: number;
  text: string;
}

/**
 * Build an edit that updates (or inserts) the `count` attribute of a container's
 * opening tag. The edit targets ONLY the attribute range, never the whole tag,
 * so it can never collide with another edit anchored at the tag's start
 * (e.g. inserting a `<numFmts>` container before `<fonts>`).
 */
function countEdit(xml: string, node: XmlNode, count: number): Edit {
  const open = xml.slice(node.start, node.openEnd);
  const match = /\bcount="\d+"/.exec(open);
  if (match) {
    const start = node.start + match.index;
    return { start, end: start + match[0]!.length, text: `count="${count}"` };
  }
  return { start: node.openEnd - 1, end: node.openEnd, text: ` count="${count}">` };
}

/**
 * Tracks and extends `xl/styles.xml` so cells can carry number formats, fonts
 * and fills. Only new entries are appended; existing ones are reused (dedup).
 */
export class StyleRegistry {
  private numFmtsNode: XmlNode | null;
  private fontsNode: XmlNode;
  private fillsNode: XmlNode;
  private bordersNode: XmlNode | null;
  private cellStyleXfsNode: XmlNode | null;
  private cellXfsNode: XmlNode;

  private numFmtByCode = new Map<string, number>();
  private numFmtById = new Map<number, string>();
  private nextNumFmtId: number;

  private fontRaws: string[] = [];
  private fillRaws: string[] = [];
  /** Canonical border XML per `borders` index (existing + added). */
  private borderRaws: string[] = [];
  /** Parsed border props per `borders` index (existing + added). */
  private borderPropsList: Array<BorderProps | undefined> = [];
  private xfDescs: XfDesc[] = [];

  private addedNumFmts: { id: number; code: string }[] = [];
  private addedFonts: string[] = [];
  private addedFills: string[] = [];
  private addedBorders: string[] = [];
  private addedXfs: string[] = [];

  constructor(private xml: string) {
    const document = parseXml(xml);
    const styleSheet = document.children.find((n) => n.name === 'styleSheet') ?? document;

    this.numFmtsNode = child(styleSheet, 'numFmts') ?? null;
    if (this.numFmtsNode) {
      for (const nf of childrenOf(this.numFmtsNode, 'numFmt')) {
        const id = Number(nf.attrs.numFmtId);
        const code = nf.attrs.formatCode ?? '';
        this.numFmtByCode.set(code, id);
        this.numFmtById.set(id, code);
      }
    }
    const maxId = [...this.numFmtById.keys()].reduce((a, b) => Math.max(a, b), 163);
    this.nextNumFmtId = Math.max(164, maxId + 1);

    const fontsNode = child(styleSheet, 'fonts');
    const fillsNode = child(styleSheet, 'fills');
    const cellXfsNode = child(styleSheet, 'cellXfs');
    if (!fontsNode || !fillsNode || !cellXfsNode) {
      throw new SheetLiteError(
        'styles.xml is missing required sections (fonts/fills/cellXfs)',
        'BAD_STYLES',
      );
    }
    this.fontsNode = fontsNode;
    this.fillsNode = fillsNode;
    this.cellXfsNode = cellXfsNode;
    this.bordersNode = child(styleSheet, 'borders') ?? null;
    this.cellStyleXfsNode = child(styleSheet, 'cellStyleXfs') ?? null;

    for (const font of childrenOf(fontsNode, 'font')) this.fontRaws.push(normalize(slice(xml, font)));
    for (const fill of childrenOf(fillsNode, 'fill')) this.fillRaws.push(normalize(slice(xml, fill)));
    if (this.bordersNode) {
      for (const border of childrenOf(this.bordersNode, 'border')) {
        this.borderRaws.push(normalize(slice(xml, border)));
        this.borderPropsList.push(parseBorderNode(border));
      }
    }
    for (const xf of childrenOf(cellXfsNode, 'xf')) {
      const alignmentNode = child(xf, 'alignment');
      const alignmentProps = alignmentNode ? parseAlignmentNode(alignmentNode) : undefined;
      this.xfDescs.push({
        numFmtId: Number(xf.attrs.numFmtId ?? '0'),
        fontId: Number(xf.attrs.fontId ?? '0'),
        fillId: Number(xf.attrs.fillId ?? '0'),
        borderId: Number(xf.attrs.borderId ?? '0'),
        alignmentProps,
        alignmentXml: alignmentXmlFor(alignmentProps),
        raw: slice(xml, xf),
      });
    }
  }

  get changed(): boolean {
    return (
      this.addedNumFmts.length +
        this.addedFonts.length +
        this.addedFills.length +
        this.addedBorders.length +
        this.addedXfs.length >
      0
    );
  }

  numFmtCode(styleIndex: number): string {
    const id = this.xfDescs[styleIndex]?.numFmtId ?? 0;
    return BUILTIN_BY_ID.get(id) ?? this.numFmtById.get(id) ?? 'General';
  }

  fontProps(styleIndex: number): FontProps | undefined {
    const xf = this.xfDescs[styleIndex];
    if (!xf) return undefined;
    const raw = this.fontRaws[xf.fontId];
    if (!raw) return undefined;
    const props: FontProps = {};
    if (/<b\b/.test(raw)) props.bold = true;
    if (/<i\b/.test(raw)) props.italic = true;
    if (/<u\b/.test(raw)) props.underline = true;
    const sz = /<sz val="([\d.]+)"/.exec(raw);
    if (sz) props.size = Number(sz[1]);
    const name = /<name val="([^"]*)"/.exec(raw);
    if (name) props.name = decodeEntities(name[1]!);
    const color = /<color rgb="([0-9A-Fa-f]{8})"/.exec(raw);
    if (color) props.color = color[1];
    return props;
  }

  fillProps(styleIndex: number): FillProps | undefined {
    const xf = this.xfDescs[styleIndex];
    if (!xf) return undefined;
    const raw = this.fillRaws[xf.fillId];
    if (!raw) return undefined;
    const pattern = /<patternFill[^>]*patternType="([^"]*)"/.exec(raw);
    if (!pattern) return undefined;
    const fg = /<fgColor rgb="([0-9A-Fa-f]{8})"/.exec(raw);
    return { type: 'pattern', pattern: decodeEntities(pattern[1]!), fgColor: fg?.[1] };
  }

  /** Border props of a `cellXfs` entry (sides with a style/color only). */
  borderProps(styleIndex: number): BorderProps | undefined {
    const xf = this.xfDescs[styleIndex];
    if (!xf) return undefined;
    const props = this.borderPropsList[xf.borderId];
    return props && Object.keys(props).length > 0 ? { ...props } : undefined;
  }

  /** Alignment of a `cellXfs` entry (from its `<alignment/>` child). */
  alignmentProps(styleIndex: number): AlignmentProps | undefined {
    const props = this.xfDescs[styleIndex]?.alignmentProps;
    return props ? { ...props } : undefined;
  }

  /**
   * Resolve a cell's final `cellXfs` index given style overrides.
   *
   * @throws {SheetLiteError} `BAD_STYLE` for a value the schema forbids (the
   *   same rule the incremental writer applies, `writer/style-validate.ts`).
   */
  resolveIndex(baseIndex: number, overrides: CellStyleOverrides): number {
    assertStylePatch(overrides);
    const base = this.xfDescs[baseIndex] ?? {
      numFmtId: 0,
      fontId: 0,
      fillId: 0,
      borderId: 0,
      alignmentXml: '',
      raw: '',
    };
    let { numFmtId, fontId, fillId, borderId } = base;
    let changed = false;

    if (overrides.numFmt !== undefined) {
      numFmtId = this.numFmtIdFor(overrides.numFmt);
      changed = true;
    }
    if (overrides.font) {
      fontId = this.fontIdFor(overrides.font, base.fontId);
      changed = true;
    }
    if (overrides.fill) {
      fillId = this.fillIdFor(overrides.fill);
      changed = true;
    }
    // Borders and alignment live per-xf, so the base xf's values are merged
    // (never dropped): adding one side or one alignment prop keeps the others.
    const baseBorder = this.borderPropsList[base.borderId];
    if (overrides.border && Object.keys(overrides.border).length > 0) {
      borderId = this.borderIdFor({ ...baseBorder, ...overrides.border }, base.borderId);
      changed = true;
    }
    const mergedAlignment = overrides.alignment
      ? { ...base.alignmentProps, ...overrides.alignment }
      : base.alignmentProps;
    if (overrides.alignment) changed = true;
    if (!changed) return baseIndex;

    // Two xfs are the same style only when ids AND alignment match.
    const alignmentXml = alignmentXmlFor(mergedAlignment);
    const existing = this.xfDescs.findIndex(
      (xf) =>
        xf.numFmtId === numFmtId &&
        xf.fontId === fontId &&
        xf.fillId === fillId &&
        xf.borderId === borderId &&
        xf.alignmentXml === alignmentXml,
    );
    if (existing !== -1) return existing;

    const apply =
      (overrides.numFmt !== undefined ? ' applyNumberFormat="1"' : '') +
      (overrides.font ? ' applyFont="1"' : '') +
      (overrides.fill ? ' applyFill="1"' : '') +
      (overrides.border && Object.keys(overrides.border).length > 0 ? ' applyBorder="1"' : '') +
      (overrides.alignment ? ' applyAlignment="1"' : '');
    const attrs =
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${apply}`;
    const raw = alignmentXml ? `${attrs}>${alignmentXml}</xf>` : `${attrs}/>`;
    this.addedXfs.push(raw);
    this.xfDescs.push({
      numFmtId,
      fontId,
      fillId,
      borderId,
      alignmentProps: Object.keys(mergedAlignment ?? {}).length > 0 ? mergedAlignment : undefined,
      alignmentXml,
      raw,
    });
    return this.xfDescs.length - 1;
  }

  toXml(): string {
    if (!this.changed) return this.xml;

    const edits: Edit[] = [];

    // New number formats (create the container if missing).
    if (this.addedNumFmts.length > 0) {
      const entries = this.addedNumFmts
        .map((nf) => `<numFmt numFmtId="${nf.id}" formatCode="${escapeXml(nf.code)}"/>`)
        .join('');
      if (this.numFmtsNode) {
        // numFmtById already contains the added formats.
        edits.push(countEdit(this.xml, this.numFmtsNode, this.numFmtById.size));
        edits.push({ start: this.numFmtsNode.closeStart, end: this.numFmtsNode.closeStart, text: entries });
      } else {
        // Inserted before <fonts>. The fonts count edit below targets ONLY its
        // count attribute, so these two edits can never collide.
        edits.push({
          start: this.fontsNode.start,
          end: this.fontsNode.start,
          text: `<numFmts count="${this.numFmtById.size}">${entries}</numFmts>`,
        });
      }
    }

    if (this.addedFonts.length > 0) {
      // fontRaws already contains the added fonts.
      edits.push(countEdit(this.xml, this.fontsNode, this.fontRaws.length));
      edits.push({ start: this.fontsNode.closeStart, end: this.fontsNode.closeStart, text: this.addedFonts.join('') });
    }

    if (this.addedFills.length > 0) {
      edits.push(countEdit(this.xml, this.fillsNode, this.fillRaws.length));
      edits.push({ start: this.fillsNode.closeStart, end: this.fillsNode.closeStart, text: this.addedFills.join('') });
    }

    if (this.addedBorders.length > 0) {
      if (this.bordersNode) {
        // borderRaws already contains the added borders.
        edits.push(countEdit(this.xml, this.bordersNode, this.borderRaws.length));
        edits.push({
          start: this.bordersNode.closeStart,
          end: this.bordersNode.closeStart,
          text: this.addedBorders.join(''),
        });
      } else {
        // Container missing entirely: create it at its schema position
        // (after <fills>, before <cellStyleXfs>/<cellXfs>).
        const anchor = this.cellStyleXfsNode ?? this.cellXfsNode;
        edits.push({
          start: anchor.start,
          end: anchor.start,
          text: `<borders count="${this.borderRaws.length}">${this.borderRaws.join('')}</borders>`,
        });
      }
    }

    if (this.addedXfs.length > 0) {
      edits.push(countEdit(this.xml, this.cellXfsNode, this.xfDescs.length));
      edits.push({
        start: this.cellXfsNode.closeStart,
        end: this.cellXfsNode.closeStart,
        text: this.addedXfs.join(''),
      });
    }

    edits.sort((a, b) => b.start - a.start);
    let out = this.xml;
    for (const edit of edits) out = out.slice(0, edit.start) + edit.text + out.slice(edit.end);
    return out;
  }

  private numFmtIdFor(code: string): number {
    const builtin = BUILTIN_NUMFMT[code];
    if (builtin !== undefined) return builtin;
    const existing = this.numFmtByCode.get(code);
    if (existing !== undefined) return existing;
    const id = this.nextNumFmtId++;
    this.numFmtByCode.set(code, id);
    this.numFmtById.set(id, code);
    this.addedNumFmts.push({ id, code });
    return id;
  }

  private fontIdFor(props: FontProps, baseFontId: number): number {
    const inner =
      (props.bold ? '<b/>' : '') +
      (props.italic ? '<i/>' : '') +
      (props.underline ? '<u/>' : '') +
      (props.size !== undefined ? `<sz val="${props.size}"/>` : '') +
      (props.color ? `<color rgb="${escapeXml(props.color)}"/>` : '') +
      (props.name ? `<name val="${escapeXml(props.name)}"/>` : '');
    if (inner === '') return baseFontId;
    const raw = `<font>${inner}</font>`;
    const existing = this.fontRaws.indexOf(raw);
    if (existing !== -1) return existing;
    this.addedFonts.push(raw);
    this.fontRaws.push(raw);
    return this.fontRaws.length - 1;
  }

  private fillIdFor(props: FillProps): number {
    const inner =
      `<patternFill patternType="${escapeXml(props.pattern ?? 'solid')}">` +
      (props.fgColor ? `<fgColor rgb="${escapeXml(props.fgColor)}"/>` : '') +
      (props.bgColor ? `<bgColor rgb="${escapeXml(props.bgColor)}"/>` : '<bgColor indexed="64"/>') +
      '</patternFill>';
    const raw = `<fill>${inner}</fill>`;
    const existing = this.fillRaws.indexOf(raw);
    if (existing !== -1) return existing;
    this.addedFills.push(raw);
    this.fillRaws.push(raw);
    return this.fillRaws.length - 1;
  }

  private borderIdFor(props: BorderProps, baseBorderId: number): number {
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
    const diagAttrs =
      (props.diagonalUp ? ' diagonalUp="1"' : '') + (props.diagonalDown ? ' diagonalDown="1"' : '');
    const raw =
      `<border${diagAttrs}>` +
      BORDER_SIDES.map((name) => side(name, props[name])).join('') +
      '</border>';
    if (raw === '<border><left/><right/><top/><bottom/><diagonal/></border>') return baseBorderId;
    const existing = this.borderRaws.indexOf(raw);
    if (existing !== -1) return existing;
    this.addedBorders.push(raw);
    this.borderRaws.push(raw);
    this.borderPropsList.push({ ...props });
    return this.borderRaws.length - 1;
  }
}

/**
 * `<alignment/>` child of an `xf`, or '' when the xf carries none. Module-level
 * (not a method) because the dedup key is built from it, for parsed and added
 * xfs alike.
 */
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
