import { strFromU8 } from 'fflate';

export interface ContentTypes {
  /** extension (lowercase, no dot) -> content type */
  defaults: Map<string, string>;
  /** absolute part name ("/xl/workbook.xml") -> content type */
  overrides: Map<string, string>;
}

/**
 * Minimal, dependency-free parser for `[Content_Types].xml`.
 * Attribute order independent; tolerant of whitespace.
 */
export function parseContentTypes(source: string | Uint8Array): ContentTypes {
  const xml = typeof source === 'string' ? source : strFromU8(source);
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();

  const defaultRe = /<Default\b[^>]*\bExtension="([^"]+)"[^>]*\bContentType="([^"]+)"[^>]*\/?>/g;
  const overrideRe = /<Override\b[^>]*\bPartName="([^"]+)"[^>]*\bContentType="([^"]+)"[^>]*\/?>/g;

  for (const match of xml.matchAll(defaultRe)) {
    defaults.set(match[1]!.toLowerCase(), match[2]!);
  }
  for (const match of xml.matchAll(overrideRe)) {
    overrides.set(match[1]!, match[2]!);
  }

  return { defaults, overrides };
}

export function contentTypeOf(ct: ContentTypes, partName: string): string | undefined {
  const override = ct.overrides.get(partName);
  if (override) return override;
  const dot = partName.lastIndexOf('.');
  if (dot < 0) return undefined;
  const ext = partName.slice(dot + 1).toLowerCase();
  return ct.defaults.get(ext);
}
