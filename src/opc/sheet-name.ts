/**
 * Worksheet-name rules Excel enforces.
 *
 * One rule, two writers: `Workbook.addWorksheet` (eager) and the incremental
 * writer both validate through here, so a name accepted by one can never be
 * rejected — or silently "repaired" by Excel — by the other.
 */
import { SheetLiteError } from '../zip/errors';
import { illegalXmlChar } from '../xml/escape';

/**
 * Reject worksheet names Excel would refuse: empty, longer than 31 characters,
 * containing `: \ / ? * [ ]`, or containing a character XML forbids (the same
 * rule every written string obeys, see `xml/escape`).
 *
 * @throws {SheetLiteError} `BAD_SHEET_NAME`.
 */
export function validateSheetName(name: string): void {
  if (typeof name !== 'string' || name.length === 0 || name.length > 31) {
    throw new SheetLiteError(
      `Invalid worksheet name ${JSON.stringify(name)}: must be 1-31 characters`,
      'BAD_SHEET_NAME',
    );
  }
  if (/[:\\/?*[\]]/.test(name)) {
    throw new SheetLiteError(
      `Invalid worksheet name ${JSON.stringify(name)}: the characters : \\ / ? * [ ] are not allowed`,
      'BAD_SHEET_NAME',
    );
  }
  // The name is interpolated into <sheet name="…">: a C0 control or a
  // non-character would make Excel refuse the whole package.
  const illegal = illegalXmlChar(name);
  if (illegal) {
    throw new SheetLiteError(
      `Invalid worksheet name ${JSON.stringify(name)}: it contains ${illegal}, which is not valid in XML`,
      'BAD_SHEET_NAME',
    );
  }
}