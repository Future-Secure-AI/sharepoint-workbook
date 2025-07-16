/**
 * Options for operation.
 * @module Options
 * @category Models
 */
import type { WorksheetName } from "./Worksheet.ts";

/**
 * Options for reading a workbook file.
 * @property {WorkbookWorksheetName} [defaultWorksheetName] Default worksheet name to use when importing a CSV file.
 * @property {(bytes: number): void} [progress] Progress callback, receives bytes processed.
 */
export type ReadOptions = {
	defaultWorksheetName?: WorksheetName;
	progress?: (bytes: number) => void;
};

/**
 * Options for writing a workbook file.
 * @property {"fail" | "replace" | "rename"} [ifExists] Behavior if the file already exists.
 * @property {(bytes: number): void} [progress] Progress callback, receives bytes processed.
 * @property {number} [maxChunkSize] Maximum chunk size in bytes for writing.
 */
export type WriteOptions = {
	ifExists?: "fail" | "replace" | "rename";
	progress?: (bytes: number) => void;
	maxChunkSize?: number;
};
