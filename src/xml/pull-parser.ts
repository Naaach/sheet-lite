/**
 * A tiny, dependency-free **incremental** XML pull-parser.
 *
 * `xml/xml.ts` builds a DOM (fine for small parts like workbook.xml or
 * styles.xml). For large parts — worksheets, shared strings — that DOM can be
 * tens or hundreds of times the size of the XML. `XmlPullParser` instead lets
 * a consumer feed text chunks and pull one event at a time, so only the
 * current element's content is ever held.
 *
 * It understands just enough XML for OOXML: elements, attributes, text runs,
 * comments, CDATA, processing instructions and doctypes. Text runs are only
 * emitted once complete (up to the next `<`), so entities can never be split
 * across events; decoding is left to the consumer.
 */
import { parseAttrs } from './xml';

export type XmlEvent =
  | { kind: 'open'; name: string; attrs: Record<string, string>; selfClosing: boolean }
  | { kind: 'close'; name: string }
  | { kind: 'text'; text: string };

const stripNs = (name: string): string => {
  const i = name.indexOf(':');
  return i === -1 ? name : name.slice(i + 1);
};

export class XmlPullParser {
  private buf = '';
  private pos = 0;
  private finished = false;

  push(chunk: string): void {
    this.buf += chunk;
  }

  end(): void {
    this.finished = true;
  }

  get ended(): boolean {
    return this.finished;
  }

  /** Characters of unconsumed input currently buffered (bounded-memory diagnostics). */
  get buffered(): number {
    return this.buf.length - this.pos;
  }

  /**
   * Return the next event, or `null` when more input is needed (or the input
   * ended).
   */
  next(): XmlEvent | null {
    // Compact the buffer occasionally to keep slicing cheap. `slice` may return
    // a V8 SlicedString referencing the parent; the parents here are only ~one
    // chunk long because compaction runs before they can grow.
    if (this.pos > 512 && this.pos * 2 > this.buf.length) {
      this.buf = this.buf.slice(this.pos);
      this.pos = 0;
    }

    for (;;) {
      const buf = this.buf;
      const start = this.pos;
      if (start >= buf.length) return null;

      // --- text run up to the next '<' ---
      if (buf.charCodeAt(start) !== 0x3c) {
        const lt = buf.indexOf('<', start);
        if (lt === -1) {
          if (!this.finished) return null; // incomplete text run: wait for more
          this.pos = buf.length;
          return { kind: 'text', text: buf.slice(start) };
        }
        if (lt > start) {
          this.pos = lt;
          return { kind: 'text', text: buf.slice(start, lt) };
        }
      }

      // --- at '<': skip comments / CDATA / PI / doctype ---
      if (buf.startsWith('<!--', start)) {
        const close = buf.indexOf('-->', start + 4);
        if (close === -1) return this.finished ? this.unterminated() : null;
        this.pos = close + 3;
        continue;
      }
      if (buf.startsWith('<![CDATA[', start)) {
        const close = buf.indexOf(']]>', start + 9);
        if (close === -1) return this.finished ? this.unterminated() : null;
        this.pos = close + 3;
        return { kind: 'text', text: buf.slice(start + 9, close) };
      }
      if (buf.startsWith('<?', start)) {
        const close = buf.indexOf('?>', start + 2);
        if (close === -1) return this.finished ? this.unterminated() : null;
        this.pos = close + 2;
        continue;
      }
      if (buf.startsWith('<!', start)) {
        const close = buf.indexOf('>', start + 2);
        if (close === -1) return this.finished ? this.unterminated() : null;
        this.pos = close + 1;
        continue;
      }

      // --- a real tag: need its '>' ---
      const gt = buf.indexOf('>', start + 1);
      if (gt === -1) return this.finished ? this.unterminated() : null;
      const raw = buf.slice(start, gt + 1);
      this.pos = gt + 1;

      if (raw.startsWith('</')) {
        return { kind: 'close', name: stripNs(raw.slice(2, -1).trim()) };
      }
      const selfClosing = raw.endsWith('/>');
      const inner = raw.slice(1, selfClosing ? -2 : -1).trim();
      const sep = inner.search(/\s/);
      const name = stripNs(sep === -1 ? inner : inner.slice(0, sep));
      const attrSource = sep === -1 ? '' : inner.slice(sep);
      return { kind: 'open', name, attrs: parseAttrs(attrSource), selfClosing };
    }
  }

  private unterminated(): never {
    throw new Error('Malformed XML: unterminated construct at end of stream');
  }
}
