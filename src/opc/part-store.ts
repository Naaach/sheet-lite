import type { ZipArchive, ZipEntry } from '../zip/archive';

/**
 * In-memory view of an OPC package: every part plus a dirty set so that a
 * write can be a true patch (only re-serialize what actually changed).
 */
export class PartStore {
  private entries: ZipEntry[];
  private byName: Map<string, ZipEntry>;
  private dirty = new Set<string>();

  /** Original bytes, used to short-circuit writes when nothing changed. */
  readonly raw: Uint8Array;

  constructor(archive: ZipArchive) {
    this.raw = archive.raw;
    this.entries = archive.entries.map((entry) => ({ ...entry }));
    this.byName = new Map(this.entries.map((entry) => [entry.name, entry]));
  }

  has(name: string): boolean {
    return this.byName.has(name);
  }

  get(name: string): Uint8Array | undefined {
    return this.byName.get(name)?.data;
  }

  names(): string[] {
    return this.entries.map((entry) => entry.name);
  }

  set(name: string, data: Uint8Array): void {
    const existing = this.byName.get(name);
    if (existing) {
      existing.data = data;
    } else {
      const entry: ZipEntry = { name, data };
      this.entries.push(entry);
      this.byName.set(name, entry);
    }
    this.dirty.add(name);
  }

  remove(name: string): void {
    if (!this.byName.has(name)) return;
    this.byName.delete(name);
    this.entries = this.entries.filter((entry) => entry.name !== name);
    this.dirty.add(name);
  }

  isDirty(): boolean {
    return this.dirty.size > 0;
  }

  dirtyNames(): string[] {
    return [...this.dirty];
  }

  toArchive(): ZipArchive {
    return { entries: this.entries, byName: this.byName, raw: this.raw };
  }
}
