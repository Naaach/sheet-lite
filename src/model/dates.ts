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

export function dateToSerial(date: Date, date1904: boolean): number {
  let serial = date.getTime() / MS_PER_DAY + EPOCH_1900;
  if (date1904) return serial - DAYS_1904_OFFSET;
  if (serial < 61) serial -= 1;
  return serial;
}

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
