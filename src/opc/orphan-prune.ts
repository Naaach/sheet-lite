import { strFromU8, strToU8 } from 'fflate';
import type { PartStore } from './part-store';
import { parseRelationships, resolvePartPath } from './relationships';
import { relsPathForPart, removeDefault, removeOverride } from './package-edit';

/**
 * Orphan-part pruning for `Workbook.removeWorksheet`.
 *
 * A worksheet references its own satellite parts through its `.rels` file
 * (drawing, legacy VML, comments, table, printer settings, hyperlinks), and
 * those parts reference more parts of their own (a drawing owns charts,
 * charts own embedded workbooks, drawings own images). When the sheet goes
 * away, everything it can reach dies with it — but only if nothing that
 * survives still points at it.
 *
 * The walk is deliberately one-sided: a part is dropped only when we can
 * PROVE that no surviving relationship targets it. A part we cannot classify
 * (unreadable rels, unknown layout) is kept: a harmless orphan in the package
 * is invisible to Excel, a dangling target is a repair prompt.
 */

const partKeyOf = (absoluteName: string): string => absoluteName.replace(/^\//, '');

/** `xl/drawings/_rels/drawing1.xml.rels` -> `xl/drawings/drawing1.xml` (undefined for an odd layout). */
function ownerOfRels(relsKey: string): string | undefined {
  const slash = relsKey.lastIndexOf('/_rels/');
  if (slash < 0) return relsKey === '_rels/.rels' ? '/' : undefined;
  return `/${relsKey.slice(0, slash)}/${relsKey.slice(slash + '/_rels/'.length, -'.rels'.length)}`;
}

/** Every `.rels` part currently in the package, including the root `_rels/.rels`. */
function relsKeys(store: PartStore): string[] {
  return store.names().filter((name) => name.includes('_rels') && name.endsWith('.rels'));
}

/**
 * The parts reachable from `rootPartKey` by following relationships, and the
 * `.rels` of each part reached: the subtree the removed sheet exclusively
 * could have owned.
 */
function ownedSubtree(store: PartStore, rootPartKey: string): Set<string> {
  const owned = new Set<string>([rootPartKey]);
  const queue = [rootPartKey];
  while (queue.length > 0) {
    const partKey = queue.pop()!;
    const relsKey = relsPathForPart(partKey);
    if (!store.has(relsKey)) continue;
    for (const rel of parseRelationships(strFromU8(store.get(relsKey)!))) {
      if (rel.external) continue;
      const target = partKeyOf(resolvePartPath(`/${partKey}`, rel.target));
      if (owned.has(target) || !store.has(target)) continue;
      owned.add(target);
      queue.push(target);
    }
  }
  return owned;
}

/**
 * Every part that survives: everything reachable WITHOUT going through the
 * removed sheet — from the package rels, the workbook rels and every
 * surviving sheet, plus whatever those parts reference in turn. Those parts
 * are off limits, and so is everything they pull in (a drawing kept because a
 * second sheet uses it keeps its charts and its images too).
 */
function reachableWithout(store: PartStore, owned: Set<string>, removedPartKey: string): Set<string> {
  const keep = new Set<string>();
  const queue: string[] = [];
  const reach = (target: string): void => {
    if (target === removedPartKey || keep.has(target)) return;
    keep.add(target);
    queue.push(target);
  };

  // Seed with the targets declared by every rels part that does NOT live inside
  // the owned subtree — a rels part that is going away protects nothing.
  for (const relsKey of relsKeys(store)) {
    const owner = ownerOfRels(relsKey);
    if (owner !== undefined && owner !== '/' && owned.has(partKeyOf(owner))) continue;
    const base = owner ?? '/';
    for (const rel of parseRelationships(strFromU8(store.get(relsKey)!))) {
      if (rel.external) continue;
      reach(partKeyOf(resolvePartPath(base, rel.target)));
    }
  }

  // Close over the rels of everything that survives.
  while (queue.length > 0) {
    const partKey = queue.pop()!;
    const relsKey = relsPathForPart(partKey);
    if (!store.has(relsKey)) continue;
    for (const rel of parseRelationships(strFromU8(store.get(relsKey)!))) {
      if (rel.external) continue;
      reach(partKeyOf(resolvePartPath(`/${partKey}`, rel.target)));
    }
  }
  return keep;
}

/** Extensions left unused once `dropped` is gone — their `Default` can go too. */
function unusedExtensions(store: PartStore, dropped: Set<string>): string[] {
  const used = new Set<string>();
  for (const name of store.names()) {
    const dot = name.lastIndexOf('.');
    if (dot > name.lastIndexOf('/')) used.add(name.slice(dot + 1).toLowerCase());
  }
  const out: string[] = [];
  for (const partKey of dropped) {
    const dot = partKey.lastIndexOf('.');
    if (dot < 0 || partKey.endsWith('.rels')) continue;
    const ext = partKey.slice(dot + 1).toLowerCase();
    if (!used.has(ext) && !out.includes(ext)) out.push(ext);
  }
  return out;
}

/**
 * Drop the parts only `sheetPartKey` owned, plus their `.rels` files and their
 * `[Content_Types].xml` entries. Must be called while the sheet is still
 * registered (part + rels present), so its own rels can be walked; the sheet
 * itself is never a drop candidate.
 *
 * Returns the part keys actually removed (empty when the sheet owned nothing).
 */
export function pruneOrphanParts(store: PartStore, sheetPartKey: string): string[] {
  const owned = ownedSubtree(store, sheetPartKey);
  const keep = reachableWithout(store, owned, sheetPartKey);

  const dropped = new Set<string>();
  for (const partKey of owned) {
    if (partKey === sheetPartKey || keep.has(partKey)) continue;
    dropped.add(partKey);
    const relsKey = relsPathForPart(partKey);
    if (store.has(relsKey)) dropped.add(relsKey);
  }
  if (dropped.size === 0) return [];

  for (const partKey of dropped) store.remove(partKey);

  const ctKey = '[Content_Types].xml';
  if (store.has(ctKey)) {
    let ct = strFromU8(store.get(ctKey)!);
    // `.rels` parts are covered by the `rels` Default and never listed.
    for (const partKey of dropped) {
      if (!partKey.endsWith('.rels')) ct = removeOverride(ct, `/${partKey}`);
    }
    // A Default may only go when NO surviving part of that extension needs it.
    for (const ext of unusedExtensions(store, dropped)) ct = removeDefault(ct, ext);
    store.set(ctKey, strToU8(ct));
  }

  return [...dropped];
}
