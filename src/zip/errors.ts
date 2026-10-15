/**
 * Base error for sheet-lite. Every error carries a stable machine-readable `code`.
 */
export class SheetLiteError extends Error {
  readonly code: string;

  constructor(message: string, code: string) {
    super(message);
    this.name = 'SheetLiteError';
    this.code = code;
  }
}

/**
 * Thrown when an input archive or part violates a safety limit
 * (zip bomb, path traversal, too many entries, oversized part, ...).
 */
export class ZipSecurityError extends SheetLiteError {
  /** `NOT_A_ZIP` for the `NotAZipError` subclass; `ZIP_SECURITY` otherwise. */
  constructor(message: string, code = 'ZIP_SECURITY') {
    super(message, code);
    this.name = 'ZipSecurityError';
  }
}

/**
 * Thrown when the input is not a ZIP container at all: a legacy OLE2/CFB
 * `.xls`, a CSV, a PDF, an empty buffer, a truncated download.
 *
 * That is a wrong format, not an attack, so it carries its own `NOT_A_ZIP`
 * code and `ZIP_SECURITY` stays reserved for genuinely hostile archives (zip
 * bombs, path traversal, limit violations, ...). It derives from
 * `ZipSecurityError` so existing `instanceof ZipSecurityError` handling of an
 * unreadable archive keeps working — match on `code` to tell them apart.
 */
export class NotAZipError extends ZipSecurityError {
  constructor(message: string) {
    super(message, 'NOT_A_ZIP');
    this.name = 'NotAZipError';
  }
}

/**
 * Thrown when an archive uses a ZIP feature sheet-lite deliberately does not
 * support (zip64, data descriptors, encryption, multi-disk archives, ...).
 */
export class ZipUnsupportedError extends SheetLiteError {
  constructor(message: string) {
    super(message, 'ZIP_UNSUPPORTED');
    this.name = 'ZipUnsupportedError';
  }
}
