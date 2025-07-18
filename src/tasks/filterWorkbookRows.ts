/**
 * Filter out unwanted rows from a workbook.
 * @module filterWorkbookRows
 * @category Tasks
 */

import type { Workbook } from "../models/Workbook.ts";

/**
 * Options for filtering workbook rows.
 * @typedef {Object} FilterWorkbookRowsOptions
 * @property {number} [skipRows] Number of rows to skip from the top of each worksheet before filtering.
 * @property {(rows: number) : void} [progress] Optional callback to report progress, called with the number of processed rows.
 */
export type FilterWorkbookRowsOptions = {
	skipRows?: number;
	progress?: (rows: number) => void;
};

/**
 * Filter out unwanted rows from a workbook.
 * @param workbook Workbook handle to filter.
 * @param rowFilter Function to determine if a row should be included, based on cell values. Return true to include the row, or false to omit it.
 * @param options Row filter options to apply (skipRows, progress).
 * @returns A promise that resolves when the filtering is complete.
 */
export default async function filterWorkbookRows(workbook: Workbook, rowFilter: (cells: string[]) => boolean, options: FilterWorkbookRowsOptions = {}): Promise<void> {
	const skipRows = options?.skipRows ?? 0;
	const progress = options?.progress ?? (() => {});

	let lastProgressTime = 0;
	let processedRows = 0;
	for (let s = 0; s < workbook.worksheets.count; s++) {
		const worksheet = workbook.worksheets.get(s);

		// Delete skipped rows from the top (before any filtering)
		if (skipRows > 0) {
			for (let i = 0; i < skipRows; i++) {
				worksheet.cells.deleteRow(0);
			}
		}

		// Filter rows
		const maxCol = worksheet.cells.maxDataColumn;
		const rowsToDelete: number[] = [];
		for (let i = 0; i <= worksheet.cells.maxDataRow; i++) {
			processedRows++;
			const rowCells = Array.from({ length: maxCol + 1 }, (_, c) => {
				const v = worksheet.cells.get(i, c)?.value;
				return typeof v === "string" || typeof v === "number" || typeof v === "boolean" ? v : "";
			});
			if (!rowFilter(rowCells.map((cell) => String(cell ?? "")))) rowsToDelete.push(i);
			if (Date.now() - lastProgressTime >= 1000) {
				progress(processedRows);
				lastProgressTime = Date.now();
			}
		}
		// Delete rows in reverse order
		rowsToDelete.sort((a, b) => b - a).forEach((rowIdx) => worksheet.cells.deleteRow(rowIdx));
		progress(processedRows);
	}
}
