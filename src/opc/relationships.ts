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
