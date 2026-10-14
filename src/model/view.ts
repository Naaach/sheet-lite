/**
 * Worksheet view (`<sheetViews>/<sheetView>/<pane>`) properties, aligned with
 * ExcelJS's `views` model where practical. Only the first view of a sheet is
 * editable; other attributes of existing views are preserved as-is.
 */
export interface SheetViewProps {
  /** `'frozen'` freezes the top-left panes; `'normal'` clears an existing pane. */
  state?: 'normal' | 'frozen' | 'frozenSplit' | 'split';
  /** Number of frozen columns (OOXML `xSplit`). */
  xSplit?: number;
  /** Number of frozen rows (OOXML `ySplit`). */
  ySplit?: number;
  /** First visible cell below/right of the split (computed from the splits when omitted). */
  topLeftCell?: string;
  workbookViewId?: number;
}