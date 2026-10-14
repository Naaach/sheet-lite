import { parseXml, child, descendants, escapeXml, type XmlNode } from '../xml/xml';

/**
 * Surgical edits on the small OPC parts that register worksheets:
 * `[Content_Types].xml`, the workbook relationships and `xl/workbook.xml`.
 * Every function rewrites only the spans it must touch — all other bytes of
 * the part stay exactly as they were.
 */

interface XmlEdit {
  start: number;
  end: number;
  text: string;
}

function applyEdits(xml: string, edits: XmlEdit[]): string {
  let out = xml;
  for (const edit of [...edits].sort((a, b) => b.start - a.start)) {
    out = out.slice(0, edit.start) + edit.text + out.slice(edit.end);
  }
  return out;
}

const span = (node: XmlNode, text: string): XmlEdit => ({ start: node.start, end: node.end, text });
const insert = (pos: number, text: string): XmlEdit => ({ start: pos, end: pos, text });

const rootOf = (xml: string, name: string): XmlNode => {
  const doc = parseXml(xml);
  return doc.children.find((n) => n.name === name) ?? doc;
};

/** Append `text` as the last child of `root`, handling self-closing containers. */
function appendChild(xml: string, root: XmlNode, text: string): XmlEdit {
  if (root.closeStart === root.openEnd) {
    // self-closing container (<Relationships/> etc.): reopen it
    const open = xml.slice(root.start, root.openEnd);
    return span(root, `${open.slice(0, -2)}>${text}</${root.name}>`);
  }
  return insert(root.closeStart, text);
}

// --- [Content_Types].xml -----------------------------------------------------

export function addOverride(xml: string, partName: string, contentType: string): string {
  const types = rootOf(xml, 'Types');
  const entry = `<Override PartName="${escapeXml(partName)}" ContentType="${escapeXml(contentType)}"/>`;
  const existing = descendants(types, 'Override').find((n) => n.attrs.PartName === partName);
  return applyEdits(xml, [existing ? span(existing, entry) : appendChild(xml, types, entry)]);
}

export function removeOverride(xml: string, partName: string): string {
  const existing = descendants(rootOf(xml, 'Types'), 'Override').find(
    (n) => n.attrs.PartName === partName,
  );
  return existing ? applyEdits(xml, [span(existing, '')]) : xml;
}

// --- relationships -----------------------------------------------------------

export function addRelationship(xml: string, id: string, type: string, target: string): string {
  const root = rootOf(xml, 'Relationships');
  const entry = `<Relationship Id="${escapeXml(id)}" Type="${escapeXml(type)}" Target="${escapeXml(target)}"/>`;
  const existing = descendants(root, 'Relationship').find((n) => n.attrs.Id === id);
  return applyEdits(xml, [existing ? span(existing, entry) : appendChild(xml, root, entry)]);
}

export function removeRelationship(xml: string, id: string): string {
  const existing = descendants(rootOf(xml, 'Relationships'), 'Relationship').find(
    (n) => n.attrs.Id === id,
  );
  return existing ? applyEdits(xml, [span(existing, '')]) : xml;
}

/** Drop the relationship whose target ends with the given part name (e.g. calcChain). */
export function removeRelationshipByTarget(xml: string, targetSuffix: string): string {
  const existing = descendants(rootOf(xml, 'Relationships'), 'Relationship').find((n) =>
    (n.attrs.Target ?? '').endsWith(targetSuffix),
  );
  return existing ? applyEdits(xml, [span(existing, '')]) : xml;
}

export function maxRelationshipId(xml: string): number {
  let max = 0;
  for (const rel of descendants(rootOf(xml, 'Relationships'), 'Relationship')) {
    const match = /^rId(\d+)$/.exec(rel.attrs.Id ?? '');
    if (match) max = Math.max(max, Number(match[1]));
  }
  return max;
}

// --- xl/workbook.xml ---------------------------------------------------------

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function addSheetEntry(xml: string, name: string, sheetId: number, rId: string): string {
  const workbook = rootOf(xml, 'workbook');
  const sheetXml = `<sheet name="${escapeXml(name)}" sheetId="${sheetId}" r:id="${escapeXml(rId)}"/>`;
  const edits: XmlEdit[] = [];

  // `r:id` needs the relationships namespace; declare it if the file does not.
  if (!/\sxmlns:r=/.test(xml.slice(workbook.start, workbook.openEnd))) {
    edits.push(insert(workbook.openEnd - 1, ` xmlns:r="${R_NS}"`));
  }

  const sheets = child(workbook, 'sheets');
  if (sheets) {
    if (sheets.closeStart === sheets.openEnd) {
      // self-closing <sheets/>: reopen it with the new entry inside
      const open = xml.slice(sheets.start, sheets.openEnd);
      edits.push(span(sheets, `${open.slice(0, -2)}>${sheetXml}</sheets>`));
    } else {
      edits.push(insert(sheets.closeStart, sheetXml));
    }
  } else {
    // A new <sheets> container must precede definedNames/calcPr/extLst.
    const anchor = child(workbook, 'definedNames') ?? child(workbook, 'calcPr') ?? child(workbook, 'extLst');
    edits.push(insert(anchor ? anchor.start : workbook.closeStart, `<sheets>${sheetXml}</sheets>`));
  }
  return applyEdits(xml, edits);
}

export function maxSheetId(xml: string): number {
  let max = 0;
  for (const sheet of descendants(rootOf(xml, 'workbook'), 'sheet')) {
    const id = Number(sheet.attrs.sheetId);
    if (Number.isFinite(id)) max = Math.max(max, id);
  }
  return max;
}

/**
 * Remove the `<sheet>` entry with the given relationship id and repair the
 * references that shift with it: defined names bound to or pointing at the
 * removed sheet (`localSheetId` indexes the remaining sheets), and the active
 * tab in `bookViews`.
 */
export function removeSheetEntry(xml: string, rId: string, sheetName: string): string {
  const workbook = rootOf(xml, 'workbook');
  const allSheets = descendants(workbook, 'sheet');
  const sheetEl = allSheets.find((n) => n.attrs['r:id'] === rId);
  if (!sheetEl) return xml;
  const index = allSheets.indexOf(sheetEl);
  const remainingCount = allSheets.length - 1;

  const edits: XmlEdit[] = [span(sheetEl, '')];

  const definedNames = child(workbook, 'definedNames');
  if (definedNames) {
    const names = descendants(definedNames, 'definedName');
    const quoted = `'${sheetName}'`;
    const bare = new RegExp(`(?<![A-Za-z0-9_'])${escapeRegExp(sheetName)}!`);
    const removed: XmlNode[] = [];
    for (const name of names) {
      const local = name.attrs.localSheetId !== undefined ? Number(name.attrs.localSheetId) : undefined;
      const formula = xml.slice(name.openEnd, name.closeStart);
      if (local === index || formula.includes(quoted) || bare.test(formula)) {
        removed.push(name);
        continue;
      }
      if (local !== undefined && local > index) {
        // Sheets after the removed one shift one position up.
        const open = xml
          .slice(name.start, name.openEnd)
          .replace(/\slocalSheetId="\d+"/, ` localSheetId="${local - 1}"`);
        edits.push({
          start: name.start,
          end: name.end,
          text: open + formula + xml.slice(name.closeStart, name.end),
        });
      }
    }
    if (names.length > 0 && removed.length === names.length) {
      edits.push(span(definedNames, ''));
    } else {
      for (const name of removed) edits.push(span(name, ''));
    }
  }

  const workbookView = child(child(workbook, 'bookViews') ?? workbook, 'workbookView');
  const activeTab = workbookView?.attrs.activeTab !== undefined ? Number(workbookView.attrs.activeTab) : NaN;
  if (workbookView && Number.isFinite(activeTab) && activeTab >= index) {
    const next = Math.max(0, Math.min(activeTab - 1, remainingCount - 1));
    const open = xml
      .slice(workbookView.start, workbookView.openEnd)
      .replace(/\sactiveTab="\d+"/, ` activeTab="${next}"`);
    edits.push({ start: workbookView.start, end: workbookView.openEnd, text: open });
  }

  return applyEdits(xml, edits);
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** `xl/worksheets/sheet2.xml` -> `xl/worksheets/_rels/sheet2.xml.rels` */
export function relsPathForPart(partName: string): string {
  const slash = partName.lastIndexOf('/');
  return `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`;
}