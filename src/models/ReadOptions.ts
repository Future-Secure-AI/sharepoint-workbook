/**
 * Configuration for a read operation.
 * @module ReadOptions
 * @category Models
 */
import type { WorkbookWorksheetName } from "microsoft-graph/dist/cjs/models/WorkbookWorksheet";

/**
 * Options for reading a workbook file.
 * @property {WorkbookWorksheetName} [defaultWorksheetName] Default worksheet name to use when importing a CSV file.
 * @property {(bytes: number): void} [progress] Progress callback, receives bytes processed.
 */
export type ReadOptions = {
	defaultWorksheetName?: WorkbookWorksheetName;
	progress?: (bytes: number) => void;
};
