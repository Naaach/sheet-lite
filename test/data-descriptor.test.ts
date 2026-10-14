import { describe, it, expect } from 'vitest';
import { unzipSync } from 'fflate';
import { readZip, writeZip, crc32 } from '../src/index';

const encode = (s: string) => new TextEncoder().encode(s);
const decode = (b: Uint8Array) => new TextDecoder().decode(b);

/**
 * Build a ZIP whose entries use the *data descriptor* convention (general
 * purpose flag bit 3): the local header has zero CRC/sizes and the real values
 * follow the data, in a 16-byte descriptor (with the optional 0x08074b50 sig).
 * This is what streaming ZIP writers emit — and what real `.xlsx` files from
 * file-examples.com use.
 */
function buildDataDescriptorZip(entries: { name: string; content: Uint8Array }[]): Uint8Array {
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  const metas = entries.map((e) => {
    const nameBytes = encode(e.name);
    return { nameBytes, data: e.content, crc: crc32(e.content) };
  });

  let offset = 0;
  const localOffsets: number[] = [];

  for (const m of metas) {
    localOffsets.push(offset);
    const header = new Uint8Array(30 + m.nameBytes.length);
    const dv = new DataView(header.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true); // version needed
    dv.setUint16(6, 0x08, true); // flags: bit 3 (data descriptor)
    dv.setUint16(8, 0, true); // method: stored
    dv.setUint16(26, m.nameBytes.length, true);
    header.set(m.nameBytes, 30);

    const descriptor = new Uint8Array(16);
    const ddv = new DataView(descriptor.buffer);
    ddv.setUint32(0, 0x08074b50, true); // descriptor signature
    ddv.setUint32(4, m.crc, true);
    ddv.setUint32(8, m.data.length, true);
    ddv.setUint32(12, m.data.length, true);

    chunks.push(header, m.data, descriptor);
    offset += header.length + m.data.length + descriptor.length;
  }

  const cdStart = offset;
  for (let i = 0; i < metas.length; i += 1) {
    const m = metas[i]!;
    const rec = new Uint8Array(46 + m.nameBytes.length);
    const dv = new DataView(rec.buffer);
    dv.setUint32(0, 0x02014b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(6, 20, true);
    dv.setUint16(8, 0x08, true); // flags: bit 3
    dv.setUint16(10, 0, true);
    dv.setUint32(16, m.crc, true);
    dv.setUint32(20, m.data.length, true);
    dv.setUint32(24, m.data.length, true);
    dv.setUint16(28, m.nameBytes.length, true);
    dv.setUint32(42, localOffsets[i]!, true);
    rec.set(m.nameBytes, 46);
    central.push(rec);
    offset += rec.length;
  }

  const eocd = new Uint8Array(22);
  const edv = new DataView(eocd.buffer);
  edv.setUint32(0, 0x06054b50, true);
  edv.setUint16(8, metas.length, true);
  edv.setUint16(10, metas.length, true);
  edv.setUint32(12, offset - cdStart, true);
  edv.setUint32(16, cdStart, true);

  const parts = [...chunks, ...central, eocd];
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  for (const p of parts) {
    out.set(p, pos);
    pos += p.length;
  }
  return out;
}

describe('ZIP data descriptors (general purpose flag bit 3)', () => {
  it('reads entries that use data descriptors', () => {
    const zip = buildDataDescriptorZip([
      { name: 'a.txt', content: encode('hello') },
      { name: 'nested/b.txt', content: encode('world!!!') },
    ]);
    const archive = readZip(zip);
    expect(decode(archive.byName.get('a.txt')!.data)).toBe('hello');
    expect(decode(archive.byName.get('nested/b.txt')!.data)).toBe('world!!!');
  });

  it('re-serializes to a valid archive (bit 3 cleared, sizes present)', () => {
    const zip = buildDataDescriptorZip([
      { name: 'a.txt', content: encode('hello') },
      { name: 'b.txt', content: encode('world!!!') },
    ]);
    const archive = readZip(zip);
    const out = writeZip(archive);

    // Our reader and fflate both accept the rebuilt archive.
    const reread = readZip(out);
    expect(decode(reread.byName.get('a.txt')!.data)).toBe('hello');
    expect(decode(reread.byName.get('b.txt')!.data)).toBe('world!!!');
    expect(() => unzipSync(out)).not.toThrow();

    // The first local header no longer advertises a data descriptor.
    const dv = new DataView(out.buffer, out.byteOffset, out.byteLength);
    expect(dv.getUint32(0, true)).toBe(0x04034b50);
    expect(dv.getUint16(6, true) & 0x08).toBe(0);
    expect(dv.getUint32(18, true)).toBe(5); // compressed size of "hello"
  });

  it('edits one entry and keeps the others intact', () => {
    const zip = buildDataDescriptorZip([
      { name: 'a.txt', content: encode('keep me') },
      { name: 'b.txt', content: encode('change me') },
    ]);
    const archive = readZip(zip);
    archive.byName.get('b.txt')!.data = encode('CHANGED');
    const out = writeZip(archive);

    const reread = readZip(out);
    expect(decode(reread.byName.get('a.txt')!.data)).toBe('keep me');
    expect(decode(reread.byName.get('b.txt')!.data)).toBe('CHANGED');
    expect(() => unzipSync(out)).not.toThrow();
  });
});
