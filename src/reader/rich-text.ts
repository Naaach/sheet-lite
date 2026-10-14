/**
 * Rich-text (`<is>` / shared-string `<si>`) run parsing.
 *
 * The three readers that can meet rich text — the DOM inline-string parser,
 * the incremental shared-strings parser and the streaming worksheet reader —
 * all feed the SAME collector here, so they can never disagree on the result:
 * a plain string when the text has no runs, or `{ richText }` (ExcelJS shape)
 * as soon as a single `<r>` run is present.
 */
import { decodeEntities, type XmlNode } from '../xml/xml';
import type { FontProps } from '../model/style';
import type { RichTextRun, RichTextValue } from '../model/value';

const isFalse = (value: string | undefined): boolean => value === '0' || value === 'false';

/**
 * Accumulates runs/text from parser events. Feed `openT`/`text`/`closeT` for
 * `<t>` content and `openRun`/`openRPr`/`fontElement`/`closeRun` for `<r>`
 * runs; text outside a `<t>` is ignored (so pretty-printed XML never leaks
 * whitespace into a string).
 */
export class RichTextRunCollector {
  private plain = '';
  private runs: RichTextRun[] | null = null;
  private run: RichTextRun | null = null;
  private buf = '';
  private inRPr = false;
  private inT = false;
  /** `<rPh>` (phonetic furigana) depth: its text is not part of the value. */
  private skipDepth = 0;

  /** Enter/leave a subtree whose text must be ignored (`<rPh>`). */
  openSkip(): void {
    this.skipDepth += 1;
  }

  closeSkip(): void {
    this.skipDepth = Math.max(0, this.skipDepth - 1);
  }

  openT(): void {
    this.inT = true;
  }

  text(raw: string): void {
    if (this.inT && this.skipDepth === 0) this.buf += raw;
  }

  closeT(): void {
    this.flush();
    this.inT = false;
  }

  openRun(): void {
    if (this.skipDepth > 0) return;
    this.flush();
    this.run = { text: '' };
    this.inRPr = false;
  }

  closeRun(): void {
    if (this.skipDepth > 0) return;
    this.flush();
    if (this.run) {
      (this.runs ??= []).push(this.run);
      this.run = null;
    }
    this.inRPr = false;
  }

  openRPr(): void {
    if (this.skipDepth > 0) return;
    this.inRPr = true;
  }

  closeRPr(): void {
    if (this.skipDepth > 0) return;
    this.inRPr = false;
  }

  /** A child element of `<rPr>` (b/i/u/sz/color/rFont). */
  fontElement(name: string, attrs: Record<string, string>): void {
    if (!this.inRPr || !this.run) return;
    const font: FontProps = this.run.font ?? (this.run.font = {});
    switch (name) {
      case 'b':
        if (!isFalse(attrs.val)) font.bold = true;
        break;
      case 'i':
        if (!isFalse(attrs.val)) font.italic = true;
        break;
      case 'u':
        if (!isFalse(attrs.val)) font.underline = true;
        break;
      case 'sz': {
        const size = Number(attrs.val);
        if (Number.isFinite(size)) font.size = size;
        break;
      }
      case 'color':
        if (attrs.rgb) font.color = attrs.rgb;
        break;
      case 'rFont':
        if (attrs.val !== undefined) font.name = attrs.val;
        break;
      default:
        break;
    }
  }

  /** The decoded value: a plain string, or `{ richText }` when runs exist. */
  value(): string | RichTextValue {
    this.flush();
    if (!this.runs || this.runs.length === 0) return this.plain;
    return { richText: this.plain ? [{ text: this.plain }, ...this.runs] : this.runs };
  }

  /** `{ richText }` when at least one `<r>` was seen, otherwise `undefined`. */
  richText(): RichTextValue | undefined {
    const value = this.value();
    return typeof value === 'string' ? undefined : value;
  }

  private flush(): void {
    if (this.buf === '') return;
    const text = decodeEntities(this.buf);
    this.buf = '';
    if (this.run) this.run.text += text;
    else this.plain += text;
  }
}

/**
 * DOM variant used for inline strings. Returns `undefined` when the element
 * has no `<r>` runs, so the caller keeps using its plain-string path.
 */
export function parseInlineRichText(isNode: XmlNode): RichTextValue | undefined {
  if (!isNode.children.some((node) => node.name === 'r')) return undefined;
  const collector = new RichTextRunCollector();
  for (const node of isNode.children) {
    if (node.name === 'r') {
      collector.openRun();
      for (const part of node.children) {
        if (part.name === 'rPr') {
          collector.openRPr();
          for (const prop of part.children) collector.fontElement(prop.name, prop.attrs);
          collector.closeRPr();
        } else if (part.name === 't') {
          collector.openT();
          collector.text(part.text);
          collector.closeT();
        }
      }
      collector.closeRun();
    } else if (node.name === 't') {
      collector.openT();
      collector.text(node.text);
      collector.closeT();
    }
  }
  return collector.richText();
}
