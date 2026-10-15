import { parseAttrs } from '../xml/xml';

export interface RelationshipRef {
  id: string;
  type: string;
  /** Target as written, relative to the declaring part unless absolute. */
  target: string;
  /** `TargetMode="External"` (hyperlinks, remote OLE links) — not a part. */
  external: boolean;
}

/**
 * Read the `Relationship` entries of a `.rels` part.
 *
 * Regex-based (like `parseContentTypes`) and namespace-prefix tolerant, so a
 * prefixed rels still reports the targets it protects.
 */
export function parseRelationships(xml: string): RelationshipRef[] {
  const out: RelationshipRef[] = [];
  const re = /<(?:[\w.-]+:)?Relationship\b([^>]*)>/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(xml))) {
    const attrs = parseAttrs(match[1]!);
    out.push({
      id: attrs.Id ?? '',
      type: attrs.Type ?? '',
      target: attrs.Target ?? '',
      external: (attrs.TargetMode ?? '').toLowerCase() === 'external',
    });
  }
  return out;
}

/**
 * Resolve a relationship Target against the part that declares it,
 * mirroring the OPC rules. Returns an absolute part name ("/a/b.xml").
 */
export function resolvePartPath(basePartName: string, target: string): string {
  if (target.startsWith('/')) return normalize(target);
  const baseDir = basePartName.includes('/')
    ? basePartName.slice(0, basePartName.lastIndexOf('/'))
    : '';
  const segments = `${baseDir}/${target}`.split('/');
  const out: string[] = [];
  for (const segment of segments) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      out.pop();
      continue;
    }
    out.push(segment);
  }
  return `/${out.join('/')}`;
}

function normalize(path: string): string {
  const out: string[] = [];
  for (const segment of path.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') {
      out.pop();
      continue;
    }
    out.push(segment);
  }
  return `/${out.join('/')}`;
}
