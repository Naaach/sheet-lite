import type { ZipArchive, ZipEntry } from '../zip/archive';

interface PartState {
  entries: ZipEntry[];
  byName: Map<string, ZipEntry>;
  dirty: Set<string>;
  raw: Uint8Array;
}

/**
 * In-memory view of an OPC package: every part plus a dirty set so that a
 * write can be a true patch (only re-serialize what actually changed).
 */
export class PartStore {
  private state: PartState;

  constructor(archive: ZipArchive) {
    const entries = archive.entries.map((entry) => ({ ...entry }));
    this.state = {
      entries,
      byName: new Map(entries.map((entry) => [entry.name, entry])),
      dirty: new Set(),
      raw: archive.raw,
    };
  }

  /** Original bytes, used to short-circuit writes when nothing changed. */
  get raw(): Uint8Array {
    return this.state.raw;
  }

  has(name: string): boolean {
    return this.state.byName.has(name);
  }

  get(name: string): Uint8Array | undefined {
    return this.state.byName.get(name)?.data;
  }

  names(): string[] {
    return this.state.entries.map((entry) => entry.name);
  }

  set(name: string, data: Uint8Array): void {
    const existing = this.state.byName.get(name);
    if (existing) {
      existing.data = data;
    } else {
      const entry: ZipEntry = { name, data };
      this.state.entries.push(entry);
      this.state.byName.set(name, entry);
    }
    this.state.dirty.add(name);
  }

  remove(name: string): void {
    if (!this.state.byName.has(name)) return;
    this.state.byName.delete(name);
    this.state.entries = this.state.entries.filter((entry) => entry.name !== name);
    this.state.dirty.add(name);
  }

  isDirty(): boolean {
    return this.state.dirty.size > 0;
  }

  dirtyNames(): string[] {
    return [...this.state.dirty];
  }

  toArchive(): ZipArchive {
    return {
      entries: this.state.entries,
      byName: this.state.byName,
      raw: this.state.raw,
    };
  }

  /** Shallow clone (entry buffers are shared) with an independent dirty set. */
  clone(): PartStore {
    const entries = this.state.entries.map((entry) => ({ ...entry }));
    return PartStore.fromState({
      entries,
      byName: new Map(entries.map((entry) => [entry.name, entry])),
      dirty: new Set(this.state.dirty),
      raw: this.state.raw,
    });
  }

  private static fromState(state: PartState): PartStore {
    const store = Object.create(PartStore.prototype) as PartStore;
    store.state = state;
    return store;
  }
}
