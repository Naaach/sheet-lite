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
  constructor(message: string) {
    super(message, 'ZIP_SECURITY');
    this.name = 'ZipSecurityError';
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
