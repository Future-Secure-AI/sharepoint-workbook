/**
 * Filter out unwanted columns from a workbook.
 * @module filterWorkbookColumns
 * @category Tasks
 */

import type { Workbook } from "aspose.cells.node";

/**
 * Filter out unwanted columns from a workbook.
 * @param workbook Workbook handle to filter.
 * @param columnFilter Function to determine if a column should be included, based on header and index. Return true to include the column, or false to omit it.
 * @returns A promise that resolves when the filtering is complete.
 */
export default async function filterWorkbookColumns(workbook: Workbook, columnFilter: (header: string, index: number) => boolean): Promise<void> {
	for (let s = 0; s < workbook.worksheets.count; s++) {
		const worksheet = workbook.worksheets.get(s);
		const maxCol = worksheet.cells.maxDataColumn;
		const headerRowIdx = 0;
		const headerCells = Array.from({ length: maxCol + 1 }, (_, c) => {
			const v = worksheet.cells.get(headerRowIdx, c)?.value;
			return typeof v === "string" || typeof v === "number" || typeof v === "boolean" ? String(v) : "";
		});
		const colIndexes = headerCells.map((header, idx) => (columnFilter(header, idx) ? idx : -1)).filter((idx) => idx !== -1);

		const removeCols = Array.from({ length: maxCol + 1 }, (_, i) => i)
			.filter((idx) => !colIndexes.includes(idx))
			.sort((a, b) => b - a);
		for (const colIdx of removeCols) worksheet.cells.deleteColumn(colIdx);
	}
}
