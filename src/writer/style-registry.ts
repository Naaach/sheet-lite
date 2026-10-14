import { parseXml, child, childrenOf, type XmlNode } from '../xml/xml';
import type { CellStyleOverrides, FillProps, FontProps } from '../model/style';

/** Common built-in number-format codes and their ids. */
const BUILTIN_NUMFMT: Record<string, number> = {
  General: 0,
  '0': 1,
  '0.00': 2,
  '#,##0': 3,
  '#,##0.00': 4,
  '0%': 9,
  '0.00%': 10,
  '0.00E+00': 11,
  '# ?/?': 12,
  '# ??/??': 13,
  'mm-dd-yy': 14,
  'd-mmm-yy': 15,
  'd-mmm': 16,
  'mmm-yy': 17,
  'h:mm AM/PM': 18,
  'h:mm:ss AM/PM': 19,
  'h:mm': 20,
  'h:mm:ss': 21,
  'm/d/yy h:mm': 22,
  '@': 49,
};
const BUILTIN_BY_ID = new Map<number, string>(Object.entries(BUILTIN_NUMFMT).map(([c, id]) => [id, c]));

interface XfDesc {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  raw: string;
}

const normalize = (xml: string): string => xml.replace(/>\s+</g, '><').trim();

function slice(xml: string, node: XmlNode): string {
  return xml.slice(node.start, node.end);
}

function setCount(tag: string, count: number): string {
  if (/\bcount="\d+"/.test(tag)) return tag.replace(/\bcount="\d+"/, `count="${count}"`);
  return tag.replace(/>$/, ` count="${count}">`);
}

/**
 * Tracks and extends `xl/styles.xml` so cells can carry number formats, fonts
 * and fills. Only new entries are appended; existing ones are reused (dedup).
 */
export class StyleRegistry {
  private numFmtsNode: XmlNode | null;
  private fontsNode: XmlNode;
  private fillsNode: XmlNode;
  private cellXfsNode: XmlNode;

  private numFmtByCode = new Map<string, number>();
  private numFmtById = new Map<number, string>();
  private nextNumFmtId: number;

  private fontRaws: string[] = [];
  private fillRaws: string[] = [];
  private xfDescs: XfDesc[] = [];

  private addedNumFmts: { id: number; code: string }[] = [];
  private addedFonts: string[] = [];
  private addedFills: string[] = [];
  private addedXfs: string[] = [];

  constructor(private xml: string) {
    const styleSheet = parseXml(xml).children.find((n) => n.name === 'styleSheet') ?? parseXml(xml);

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

    this.fontsNode = child(styleSheet, 'fonts')!;
    for (const font of childrenOf(this.fontsNode, 'font')) this.fontRaws.push(normalize(slice(xml, font)));

    this.fillsNode = child(styleSheet, 'fills')!;
    for (const fill of childrenOf(this.fillsNode, 'fill')) this.fillRaws.push(normalize(slice(xml, fill)));

    this.cellXfsNode = child(styleSheet, 'cellXfs')!;
    for (const xf of childrenOf(this.cellXfsNode, 'xf')) {
      this.xfDescs.push({
        numFmtId: Number(xf.attrs.numFmtId ?? '0'),
        fontId: Number(xf.attrs.fontId ?? '0'),
        fillId: Number(xf.attrs.fillId ?? '0'),
        borderId: Number(xf.attrs.borderId ?? '0'),
        raw: slice(xml, xf),
      });
    }
  }

  get changed(): boolean {
    return (
      this.addedNumFmts.length + this.addedFonts.length + this.addedFills.length + this.addedXfs.length > 0
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
    if (name) props.name = name[1];
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
    return { type: 'pattern', pattern: pattern[1], fgColor: fg?.[1] };
  }

  /** Resolve a cell's final `cellXfs` index given style overrides. */
  resolveIndex(baseIndex: number, overrides: CellStyleOverrides): number {
    const base = this.xfDescs[baseIndex] ?? { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, raw: '' };
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
    if (!changed) return baseIndex;

    const existing = this.xfDescs.findIndex(
      (xf) => xf.numFmtId === numFmtId && xf.fontId === fontId && xf.fillId === fillId && xf.borderId === borderId,
    );
    if (existing !== -1) return existing;

    const apply =
      (overrides.numFmt !== undefined ? ' applyNumberFormat="1"' : '') +
      (overrides.font ? ' applyFont="1"' : '') +
      (overrides.fill ? ' applyFill="1"' : '');
    const raw = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${apply}/>`;
    this.addedXfs.push(raw);
    this.xfDescs.push({ numFmtId, fontId, fillId, borderId, raw });
    return this.xfDescs.length - 1;
  }

  toXml(): string {
    if (!this.changed) return this.xml;

    const edits: { start: number; end: number; text: string }[] = [];

    // New number formats (create the container if missing).
    if (this.addedNumFmts.length > 0) {
      const entries = this.addedNumFmts.map((nf) => `<numFmt numFmtId="${nf.id}" formatCode="${nf.code}"/>`).join('');
      if (this.numFmtsNode) {
        const count = this.numFmtById.size + this.addedNumFmts.length;
        const open = setCount(this.xml.slice(this.numFmtsNode.start, this.numFmtsNode.openEnd), count);
        edits.push({ start: this.numFmtsNode.start, end: this.numFmtsNode.openEnd, text: open });
        edits.push({ start: this.numFmtsNode.closeStart, end: this.numFmtsNode.closeStart, text: entries });
      } else {
        edits.push({
          start: this.fontsNode.start,
          end: this.fontsNode.start,
          text: `<numFmts count="${this.addedNumFmts.length}">${entries}</numFmts>`,
        });
      }
    }

    if (this.addedFonts.length > 0) {
      const count = this.fontRaws.length + this.addedFonts.length;
      const open = setCount(this.xml.slice(this.fontsNode.start, this.fontsNode.openEnd), count);
      edits.push({ start: this.fontsNode.start, end: this.fontsNode.openEnd, text: open });
      edits.push({ start: this.fontsNode.closeStart, end: this.fontsNode.closeStart, text: this.addedFonts.join('') });
    }

    if (this.addedFills.length > 0) {
      const count = this.fillRaws.length + this.addedFills.length;
      const open = setCount(this.xml.slice(this.fillsNode.start, this.fillsNode.openEnd), count);
      edits.push({ start: this.fillsNode.start, end: this.fillsNode.openEnd, text: open });
      edits.push({ start: this.fillsNode.closeStart, end: this.fillsNode.closeStart, text: this.addedFills.join('') });
    }

    if (this.addedXfs.length > 0) {
      const count = this.xfDescs.length;
      const open = setCount(this.xml.slice(this.cellXfsNode.start, this.cellXfsNode.openEnd), count);
      edits.push({ start: this.cellXfsNode.start, end: this.cellXfsNode.openEnd, text: open });
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
      (props.color ? `<color rgb="${props.color}"/>` : '') +
      (props.name ? `<name val="${props.name}"/>` : '');
    if (inner === '') return baseFontId;
    const raw = `<font>${inner}</font>`;
    const existing = this.fontRaws.indexOf(raw);
    if (existing !== -1) return existing;
    this.addedFonts.push(raw);
    this.fontRaws.push(raw);
    return this.fontRaws.length - 1;
  }

  private fillIdFor(props: FillProps): number {
    const pattern = props.pattern ?? 'solid';
    const inner =
      `<patternFill patternType="${pattern}">` +
      (props.fgColor ? `<fgColor rgb="${props.fgColor}"/>` : '') +
      (props.bgColor ? `<bgColor rgb="${props.bgColor}"/>` : '<bgColor indexed="64"/>') +
      '</patternFill>';
    const raw = `<fill>${inner}</fill>`;
    const existing = this.fillRaws.indexOf(raw);
    if (existing !== -1) return existing;
    this.addedFills.push(raw);
    this.fillRaws.push(raw);
    return this.fillRaws.length - 1;
  }
}
