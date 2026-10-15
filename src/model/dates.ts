import { SheetLiteError } from '../zip/errors';

const MS_PER_DAY = 86_400_000;

/** Excel serial number of 1970-01-01 in the 1900 date system. */
const EPOCH_1900 = 25569;
/** Days between the 1904 and 1900 date-system origins. */
const DAYS_1904_OFFSET = 1462;

export function serialToDate(serial: number, date1904: boolean): Date {
  let s = date1904 ? serial + DAYS_1904_OFFSET : serial;
  // Excel wrongly treats 1900 as a leap year: serials below the phantom
  // 1900-02-29 are off by one on the shifted scale.
  if (!date1904 && s < 60) s += 1;
  return new Date(Math.round((s - EPOCH_1900) * MS_PER_DAY));
}

/**
 * Serial number of a `Date`, in the 1900 system (default) or the 1904 one.
 *
 * The 1904 system starts at serial 0 (1904-01-01): it has no representation
 * for anything earlier, and a negative `<v>` is a file Excel offers to repair,
 * so such a date is rejected instead of being written as a wrong number. The
 * 1900 system keeps its historical behaviour for pre-1900 dates.
 *
 * @throws {SheetLiteError} `BAD_DATE` for a date the 1904 system cannot hold.
 */
export function dateToSerial(date: Date, date1904: boolean): number {
  let serial = date.getTime() / MS_PER_DAY + EPOCH_1900;
  if (date1904) {
    serial -= DAYS_1904_OFFSET;
    // `!(serial >= 0)` also catches an invalid Date (NaN).
    if (!(serial >= 0)) {
      throw new SheetLiteError(
        `Cannot write ${Number.isNaN(date.getTime()) ? 'an invalid Date' : date.toISOString()}: the 1904 date system cannot represent it (its serials start at 0 = 1904-01-01)`,
        'BAD_DATE',
      );
    }
    return serial;
  }
  if (serial < 61) serial -= 1;
  return serial;
}

/** Common built-in number-format codes and their ids. */
export const BUILTIN_NUMFMT: Record<string, number> = {
  General: 0,
  '0': 1,
  '0.00': 2,
  '#,##0': 3,
  '#,##0.00': 4,
  '0%': 9,
  '0.00%': 10,
  '0.00E+00': 11,
  '# ?/?': 12,
  '# ??/??': 13,
  'mm-dd-yy': 14,
  'd-mmm-yy': 15,
  'd-mmm': 16,
  'mmm-yy': 17,
  'h:mm AM/PM': 18,
  'h:mm:ss AM/PM': 19,
  'h:mm': 20,
  'h:mm:ss': 21,
  'm/d/yy h:mm': 22,
  '@': 49,
};

/** Reverse lookup for `BUILTIN_NUMFMT` (ids not listed there fall through). */
export const BUILTIN_BY_ID = new Map<number, string>(
  Object.entries(BUILTIN_NUMFMT).map(([code, id]) => [id, code]),
);

/** Built-in Excel number-format ids that represent dates/times. */
const BUILTIN_DATE_FORMATS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
  45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
]);

export function isDateFormatCode(code: string): boolean {
  // Remove quoted literals, bracket sections ([Red], [$€-x], [$-409]) and
  // escaped characters, then look for date/time tokens.
  const cleaned = code
    .replace(/"[^"]*"/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\\./g, '');
  return /[ymdhs]/i.test(cleaned);
}

export function isBuiltinDateFormat(numFmtId: number): boolean {
  return BUILTIN_DATE_FORMATS.has(numFmtId);
}
