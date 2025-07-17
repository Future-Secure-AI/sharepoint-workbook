/**
 * Filter out unwanted rows and columns from a workbook.
 * @module filterWorkbook
 * @category Tasks
 */

import type { Handle } from "../models/Handle.ts";

/**
 * Filter options for filtering a workbook.
 * @property skipRows Number of rows to skip from the top (e.g., header rows).
 * @property column Function to determine if a column should be included. Return true to include the column, or false to omit it.
 * @property row Function to determine if a row should be included. Return true to include the row, or false to omit it.
 */
export type Filter = {
	/** Number of rows to skip from the top (e.g., header rows). */
	skipRows?: number;
	/**
	 * Function to determine if a column should be included, based on header and index.
	 * Return true to include the column, or false to omit it.
	 */
	columnFilter?: (header: string, index: number) => boolean;
	/**
	 * Function to determine if a row should be included, based on cell values.
	 * Return true to include the row, or false to omit it.
	 */
	rowFilter?: (cells: string[]) => boolean;

	progress?: (rows: number) => void;
};

/**
 * Filter out unwanted rows and columns from a workbook. All styling is lost when filtering.
 * @param handle Workbook handle to filter.
 * @param filter Filter options to apply (skipRows, column, row).
 * @returns A promise that resolves when the filtering is complete.
 */

export default async function filterWorkbook(handle: Handle, filter: Filter): Promise<void> {
	const skipRows = filter?.skipRows ?? 0;
	const columnFilter = filter?.columnFilter ?? (() => true);
	const rowFilter = filter?.rowFilter ?? (() => true);
	const progress = filter?.progress ?? (() => {});

	const workbook = handle.workbook;

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

		// Get header row and determine columns to keep
		const maxCol = worksheet.cells.maxDataColumn;
		const headerRowIdx = 0;
		const headerCells = Array.from({ length: maxCol + 1 }, (_, c) => {
			const v = worksheet.cells.get(headerRowIdx, c)?.value;
			return typeof v === "string" || typeof v === "number" || typeof v === "boolean" ? String(v) : "";
		});
		const colIndexes = headerCells.map((header, idx) => (columnFilter(header, idx) ? idx : -1)).filter((idx) => idx !== -1);

		// Delete columns not in colIndexes (right to left)
		const removeCols = Array.from({ length: maxCol + 1 }, (_, i) => i)
			.filter((idx) => !colIndexes.includes(idx))
			.sort((a, b) => b - a);
		for (const colIdx of removeCols) worksheet.cells.deleteColumn(colIdx);

		// Filter rows
		const rowsToDelete: number[] = [];
		for (let i = 0; i <= worksheet.cells.maxDataRow; i++) {
			processedRows++;
			const rowCells = colIndexes.map((c) => {
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
