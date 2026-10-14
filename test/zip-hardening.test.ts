import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { readZip, ZipSecurityError, isSafeEntryName } from '../src/index';

describe('v0.1 — ZIP hardening', () => {
  it('accepts normal part names', () => {
    expect(isSafeEntryName('xl/worksheets/sheet1.xml')).toBe(true);
    expect(isSafeEntryName('[Content_Types].xml')).toBe(true);
  });

  it('rejects traversal, absolute paths and null bytes', () => {
    expect(isSafeEntryName('../evil.txt')).toBe(false);
    expect(isSafeEntryName('/etc/passwd')).toBe(false);
    expect(isSafeEntryName('C:/windows/system32')).toBe(false);
    expect(isSafeEntryName('a/../../b')).toBe(false);
    expect(isSafeEntryName('bad\u0000name')).toBe(false);
  });

  it('rejects an archive containing a traversal entry', () => {
    const zip = zipSync({ '../evil.txt': strToU8('x') });
    expect(() => readZip(zip)).toThrow(ZipSecurityError);
  });

  it('enforces the entry-count limit', () => {
    const files: Record<string, Uint8Array> = {};
    for (let i = 0; i < 5; i += 1) files[`f${i}.txt`] = strToU8('x');
    const zip = zipSync(files);
    expect(() => readZip(zip, { maxEntries: 2 })).toThrow(ZipSecurityError);
  });

  it('enforces the total decompressed-size limit', () => {
    const zip = zipSync({ 'a.txt': strToU8('hello world') });
    expect(() => readZip(zip, { maxTotalUncompressed: 3 })).toThrow(ZipSecurityError);
  });

  it('rejects a suspicious compression ratio', () => {
    const zip = zipSync({ 'a.txt': strToU8('a'.repeat(5000)) });
    expect(() => readZip(zip, { maxCompressionRatio: 1 })).toThrow(ZipSecurityError);
  });
});
