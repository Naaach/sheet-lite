/**
 * A tiny, dependency-free XML reader.
 * It is intentionally minimal: enough to read and patch the well-known OOXML
 * parts (workbook, relationships, shared strings, styles, worksheets) while
 * preserving every byte we do not touch.
 */

export interface XmlNode {
  /** '#document' for the synthetic root. */
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  /** Concatenated raw text directly inside this element. */
  text: string;
  /** Index of the opening '<'. */
  start: number;
  /** Index just after the opening tag's '>'. */
  openEnd: number;
  /** Index of the closing '</' (equals `end` for self-closing tags). */
  closeStart: number;
  /** Index just after the closing tag's '>'. */
  end: number;
}

const TAG_RE = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<![^>]*>|<\/[^>]*>|<[^>]*\/>|<[^>]*>/g;

function makeNode(name: string): XmlNode {
  return {
    name,
    attrs: {},
    children: [],
    text: '',
    start: 0,
    openEnd: 0,
    closeStart: 0,
    end: 0,
  };
}

export function parseAttrs(source: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const re = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(source))) {
    attrs[match[1]!] = decodeEntities(match[2] ?? match[3] ?? '');
  }
  return attrs;
}

export function parseXml(xml: string): XmlNode {
  const root = makeNode('#document');
  const stack: XmlNode[] = [root];
  let last = 0;
  const re = new RegExp(TAG_RE.source, 'g');
  let match: RegExpExecArray | null;

  while ((match = re.exec(xml))) {
    const raw = match[0];
    const start = match.index;
    const end = start + raw.length;

    if (start > last) {
      const text = xml.slice(last, start);
      const top = stack[stack.length - 1]!;
      if (text.length > 0) top.text += text;
    }
    last = end;

    // Skip comments, CDATA, processing instructions and DOCTYPE.
    if (raw.startsWith('<!--') || raw.startsWith('<![CDATA[') || raw.startsWith('<?') || raw.startsWith('<!')) {
      if (raw.startsWith('<![CDATA[')) {
        stack[stack.length - 1]!.text += raw.slice(9, -3);
      } else if (raw.startsWith('<!--')) {
        // ignore comments
      } else if (raw.startsWith('<?')) {
        // ignore processing instructions
      }
      continue;
    }

    if (raw.startsWith('</')) {
      const node = stack.pop();
      if (node) {
        node.closeStart = start;
        node.end = end;
      }
      continue;
    }

    const selfClosing = raw.endsWith('/>');
    const inner = raw.slice(1, selfClosing ? -2 : -1).trim();
    const sep = inner.search(/\s/);
    const name = sep === -1 ? inner : inner.slice(0, sep);
    const attrSource = sep === -1 ? '' : inner.slice(sep);

    const node = makeNode(name);
    node.attrs = parseAttrs(attrSource);
    node.start = start;
    node.openEnd = end;
    stack[stack.length - 1]!.children.push(node);

    if (selfClosing) {
      node.closeStart = end;
      node.end = end;
    } else {
      stack.push(node);
    }
  }

  if (last < xml.length) {
    stack[stack.length - 1]!.text += xml.slice(last);
  }

  return root;
}

export function child(node: XmlNode, name: string): XmlNode | undefined {
  return node.children.find((c) => c.name === name);
}

export function childrenOf(node: XmlNode, name: string): XmlNode[] {
  return node.children.filter((c) => c.name === name);
}

/** All descendant elements with the given name (depth-first). */
export function descendants(node: XmlNode, name: string): XmlNode[] {
  const out: XmlNode[] = [];
  const walk = (n: XmlNode): void => {
    for (const c of n.children) {
      if (c.name === name) out.push(c);
      walk(c);
    }
  };
  walk(node);
  return out;
}

/** Concatenated, entity-decoded text of an element (including nested <t>). */
export function textOf(node: XmlNode | undefined): string {
  if (!node) return '';
  let out = decodeEntities(node.text);
  for (const c of node.children) out += textOf(c);
  return out;
}

const NAMED: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

export function decodeEntities(value: string): string {
  if (!value.includes('&')) return value;
  return value.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, entity: string) => {
    const toChar = (codePoint: number): string =>
      Number.isInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff
        ? String.fromCodePoint(codePoint)
        : '\uFFFD';
    if (entity.startsWith('#x') || entity.startsWith('#X')) {
      return toChar(parseInt(entity.slice(2), 16));
    }
    if (entity.startsWith('#')) {
      return toChar(parseInt(entity.slice(1), 10));
    }
    return NAMED[entity] ?? whole;
  });
}

export function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
