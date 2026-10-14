import { describe, it, expect } from 'vitest';
import { parseContentTypes, contentTypeOf, resolvePartPath } from '../src/index';

describe('v0.1 — OPC helpers', () => {
  it('parses defaults and overrides from [Content_Types].xml', () => {
    const ct = parseContentTypes(`<Types>
      <Default Extension="rels" ContentType="rels+xml"/>
      <Default Extension="XML" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="main+xml"/>
    </Types>`);
    expect(ct.defaults.get('rels')).toBe('rels+xml');
    expect(ct.defaults.get('xml')).toBe('application/xml');
    expect(ct.overrides.get('/xl/workbook.xml')).toBe('main+xml');
  });

  it('resolves the content type of a part (override wins over default)', () => {
    const ct = parseContentTypes(`<Types>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="main+xml"/>
    </Types>`);
    expect(contentTypeOf(ct, '/xl/workbook.xml')).toBe('main+xml');
    expect(contentTypeOf(ct, '/xl/styles.xml')).toBe('application/xml');
  });

  it('resolves relationship targets relative to the declaring part', () => {
    expect(resolvePartPath('/xl/workbook.xml', 'worksheets/sheet1.xml')).toBe(
      '/xl/worksheets/sheet1.xml',
    );
    expect(resolvePartPath('/xl/worksheets/sheet1.xml', '../charts/chart1.xml')).toBe(
      '/xl/charts/chart1.xml',
    );
    expect(resolvePartPath('/xl/workbook.xml', '/docProps/core.xml')).toBe('/docProps/core.xml');
    expect(resolvePartPath('/xl/workbook.xml', '../docProps/core.xml')).toBe('/docProps/core.xml');
  });
});
