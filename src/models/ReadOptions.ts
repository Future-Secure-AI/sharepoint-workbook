import type { WorkbookWorksheetName } from "microsoft-graph/WorkbookWorksheet";

/**
 * Options for reading a workbook file.
 * @typedef {Object} ReadOptions
 * @property {WorkbookWorksheetName} [defaultWorksheetName] Default worksheet name to use when importing a CSV file.
 * @property {function(number): void} [progress] Progress callback, receives bytes processed.
 */

export type ReadOptions = {
	defaultWorksheetName?: WorkbookWorksheetName;
	progress?: (bytes: number) => void;
};
