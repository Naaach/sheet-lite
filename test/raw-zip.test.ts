import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import {
  Workbook,
  readZip,
  writeZip,
  crc32,
  ZipSecurityError,
  ZipUnsupportedError,
  type ZipArchive,
} from '../src/index';
import { makeMinimalXlsx } from './fixtures/make';

const eq = (a: Uint8Array, b: Uint8Array): boolean => Buffer.from(a).equals(Buffer.from(b));

/** Assert a call throws `ZipUnsupportedError` (code ZIP_UNSUPPORTED) matching `pattern`. */
function expectZipUnsupported(fn: () => unknown, pattern: RegExp): void {
  let caught: unknown;
  try {
    fn();
  } catch (error) {
    caught = error;
  }
  expect(caught).toBeInstanceOf(ZipUnsupportedError);
  expect((caught as ZipUnsupportedError).code).toBe('ZIP_UNSUPPORTED');
  expect((caught as Error).message).toMatch(pattern);
}

/**
 * Independent central-directory walk (does not use the code under test):
 * returns, per entry name, the full local-entry byte range
 * [localOffset .. localOffset + 30 + nameLen + extraLen + compressedSize).
 */
function localEntryRanges(zip: Uint8Array): Map<string, Uint8Array> {
  const dv = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  let eocd = zip.length - 22;
  while (dv.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  const total = dv.getUint16(eocd + 10, true);
  const ranges = new Map<string, Uint8Array>();
  let pos = dv.getUint32(eocd + 16, true);
  for (let i = 0; i < total; i += 1) {
    if (dv.getUint32(pos, true) !== 0x02014b50) throw new Error('bad fixture: corrupt central directory');
    const nameLen = dv.getUint16(pos + 28, true);
    const extraLen = dv.getUint16(pos + 30, true);
    const commentLen = dv.getUint16(pos + 32, true);
    const compressedSize = dv.getUint32(pos + 20, true);
    const localOffset = dv.getUint32(pos + 42, true);
    const name = Buffer.from(zip.subarray(pos + 46, pos + 46 + nameLen)).toString('utf8');
    if (!name.endsWith('/')) {
      const localNameLen = dv.getUint16(localOffset + 26, true);
      const localExtraLen = dv.getUint16(localOffset + 28, true);
      const end = localOffset + 30 + localNameLen + localExtraLen + compressedSize;
      ranges.set(name, zip.subarray(localOffset, end));
    }
    pos += 46 + nameLen + extraLen + commentLen;
  }
  return ranges;
}

describe('runtime — raw ZIP layer (verbatim reuse)', () => {
  it('no-op write through the raw path is byte-identical to the original file', () => {
    const original = makeMinimalXlsx();
    const out = writeZip(readZip(original));
    expect(eq(out, original)).toBe(true);
  });

  it('repeated read/write cycles stay byte-identical', () => {
    const original = makeMinimalXlsx();
    let bytes = original;
    for (let i = 0; i < 3; i += 1) bytes = writeZip(readZip(bytes));
    expect(eq(bytes, original)).toBe(true);
  });

  it('editing one cell re-serializes only that entry; every other entry stays byte-identical', () => {
    const original = makeMinimalXlsx();
    const wb = Workbook.read(original);
    wb.getWorksheet('Sheet1')!.getCell('B1').value = 100;
    const out = wb.toBuffer();

    const before = localEntryRanges(original);
    const after = localEntryRanges(out);
    expect([...after.keys()].sort()).toEqual([...before.keys()].sort());

    for (const [name, bytes] of before) {
      if (name === 'xl/worksheets/sheet1.xml') continue;
      expect(eq(bytes, after.get(name)!), `entry changed: ${name}`).toBe(true);
    }
    // ...and the edited entry is really different and parses back
    expect(eq(before.get('xl/worksheets/sheet1.xml')!, after.get('xl/worksheets/sheet1.xml')!)).toBe(false);
    expect(Workbook.read(out).getWorksheet('Sheet1')!.getCell('B1').value).toBe(100);
  });

  it('reuses original compressed bytes + CRC for untouched entries (method 8 and 0)', () => {
    // One stored (level 0) and one deflated entry in the same archive.
    const zip = zipSync(
      {
        'a.txt': [strToU8('stored content, stored content, stored content'), { level: 0 }],
        'b.txt': [strToU8('deflated content, deflated content, deflated content'), { level: 6 }],
      },
      { level: 0 }, // per-entry options decide the method
    );
    const archive = readZip(zip);
    expect(archive.byName.get('a.txt')!.source!.method).toBe(0);
    expect(archive.byName.get('b.txt')!.source!.method).toBe(8);

    // No-op: byte-identical even with mixed methods.
    expect(eq(writeZip(archive), zip)).toBe(true);

    // Edit one entry: the other's local bytes are copied verbatim.
    const edited = readZip(zip);
    edited.byName.get('b.txt')!.data = strToU8('changed');
    const out = writeZip(edited);
    const before = localEntryRanges(zip);
    const after = localEntryRanges(out);
    expect(eq(before.get('a.txt')!, after.get('a.txt')!)).toBe(true);
    expect(eq(before.get('b.txt')!, after.get('b.txt')!)).toBe(false);
    expect(new TextDecoder().decode(readZip(out).byName.get('b.txt')!.data)).toBe('changed');
  });

  it('preserves EOCD comments on write', () => {
    const base = makeMinimalXlsx();
    // Rebuild the archive with a non-empty EOCD comment.
    let eocd = base.length - 22;
    while (new DataView(base.buffer).getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
    const body = base.subarray(0, eocd);
    const comment = strToU8('hello');
    const rebuilt = new Uint8Array(body.length + 22 + comment.length);
    rebuilt.set(body, 0);
    rebuilt.set(base.subarray(eocd), body.length);
    new DataView(rebuilt.buffer).setUint16(body.length + 20, comment.length, true);
    rebuilt.set(comment, body.length + 22);

    expect(eq(writeZip(readZip(rebuilt)), rebuilt)).toBe(true);
  });

  it('rejects zip64 archives with a clear error code', () => {
    const base = makeMinimalXlsx();
    const dv = new DataView(base.buffer);
    let eocd = base.length - 22;
    while (dv.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;

    // EOCD total-entries = 0xFFFF -> zip64 marker
    const patched = new Uint8Array(base);
    new DataView(patched.buffer).setUint16(eocd + 10, 0xffff, true);
    expectZipUnsupported(() => readZip(patched), /zip64/);

    // central-directory offset = 0xFFFFFFFF -> zip64 marker
    const patched2 = new Uint8Array(base);
    new DataView(patched2.buffer).setUint32(eocd + 16, 0xffffffff, true);
    expectZipUnsupported(() => readZip(patched2), /zip64/);

    // zip64 EOCD locator right before the EOCD -> zip64 marker
    const body = base.subarray(0, base.length - 22);
    const patched3 = new Uint8Array(body.length + 20 + 22);
    patched3.set(body, 0);
    new DataView(patched3.buffer).setUint32(body.length, 0x07064b50, true);
    patched3.set(base.subarray(base.length - 22), body.length + 20);
    expectZipUnsupported(() => readZip(patched3), /locator/);
  });

  it('accepts data descriptors (general purpose flag bit 3)', () => {
    const base = makeMinimalXlsx();
    const patched = new Uint8Array(base);
    const dv = new DataView(patched.buffer);
    let eocd = base.length - 22;
    while (dv.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
    let pos = dv.getUint32(eocd + 16, true);
    const total = dv.getUint16(eocd + 10, true);
    for (let i = 0; i < total; i += 1) {
      const nameLen = dv.getUint16(pos + 28, true);
      const extraLen = dv.getUint16(pos + 30, true);
      const commentLen = dv.getUint16(pos + 32, true);
      const localOffset = dv.getUint32(pos + 42, true);
      dv.setUint16(pos + 8, dv.getUint16(pos + 8, true) | 0x0008, true); // central flags
      dv.setUint16(localOffset + 6, dv.getUint16(localOffset + 6, true) | 0x0008, true); // local flags
      pos += 46 + nameLen + extraLen + commentLen;
    }
    // Streaming ZIP writers set bit 3 routinely; the central directory still
    // carries the real CRC/sizes, so the archive is read normally.
    const archive = readZip(patched);
    expect(new TextDecoder().decode(archive.byName.get('xl/workbook.xml')!.data)).toContain('<workbook');
  });

  it('rejects encrypted entries and unknown compression methods', () => {
    const base = makeMinimalXlsx();
    const dv = new DataView(base.buffer);
    let eocd = base.length - 22;
    while (dv.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
    const firstRecord = dv.getUint32(eocd + 16, true);

    const encrypted = new Uint8Array(base);
    new DataView(encrypted.buffer).setUint16(firstRecord + 8, 0x0001, true);
    expectZipUnsupported(() => readZip(encrypted), /encrypt/i);

    const method99 = new Uint8Array(base);
    new DataView(method99.buffer).setUint16(firstRecord + 10, 99, true);
    expectZipUnsupported(() => readZip(method99), /method 99/);
  });

  it('rejects corrupted entries via CRC validation', () => {
    const payload = strToU8('hello world, hello world, hello world');
    const zip = zipSync({ 'a.txt': payload }, { level: 0 }); // stored: flip a data byte
    const dv = new DataView(zip.buffer);
    dv.setUint8(30 + 5, 88); // 'o' -> 'X' inside the stored data
    expect(() => readZip(zip)).toThrow(ZipSecurityError);
    expect(() => readZip(zip)).toThrow(/CRC/);
  });

  it('rejects input that is not a ZIP archive', () => {
    expect(() => readZip(strToU8('this is definitely not a zip file'))).toThrow(ZipSecurityError);
    expect(() => readZip(new Uint8Array(0))).toThrow(ZipSecurityError);
  });

  it('crc32 matches the reference implementation', () => {
    expect(crc32(strToU8('123456789'))).toBe(0xcbf43926);
    expect(crc32(new Uint8Array(0))).toBe(0);
    expect(crc32(strToU8('The quick brown fox jumps over the lazy dog'))).toBe(0x414fa339);
  });

  it('rejects unsafe entry names on write (low-level path)', () => {
    const archive: ZipArchive = {
      entries: [{ name: '../evil.txt', data: strToU8('x') }],
      byName: new Map(),
      raw: new Uint8Array(0),
    };
    expect(() => writeZip(archive)).toThrow(ZipSecurityError);
  });

  it('round-trips a workbook written through the raw layer', () => {
    const original = makeMinimalXlsx();
    const wb = Workbook.read(original);
    wb.getWorksheet('Sheet1')!.getCell('A2').value = 'Mundo';
    const out = wb.toBuffer();
    const reread = Workbook.read(out);
    expect(reread.getWorksheet('Sheet1')!.getCell('A2').value).toBe('Mundo');
    expect(reread.getWorksheet('Sheet1')!.getCell('B1').value).toBe(42);
  });
});
