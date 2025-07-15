/**
 * Filter out unwanted rows and columns from a workbook.
 * @module filterWorkbook
 * @category Tasks
 */

import ExcelJS, { type CellValue } from "exceljs";
import type { Handle } from "../models/Handle.ts";
import { getLatestRevisionFilePath, getNextRevisionFilePath } from "../services/workingFolder.ts";

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
};

/**
 * Filter out unwanted rows and columns from a workbook. All styling is lost when filtering.
 * @param handle Workbook handle to filter.
 * @param filter Filter options to apply (skipRows, column, row).
 * @returns A promise that resolves when the filtering is complete.
 */
export default async function filterWorkbook(handle: Handle, filter: Filter): Promise<void> {
	const { skipRows = 0, columnFilter = () => true, rowFilter = () => true } = filter;
	const latestFile = await getLatestRevisionFilePath(handle.id);
	const nextFile = await getNextRevisionFilePath(handle.id);

	const reader = new ExcelJS.stream.xlsx.WorkbookReader(latestFile, { entries: "emit" });
	const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: nextFile });

	for await (const sheet of reader) {
		const sheetName = (sheet as { name?: string }).name ?? "Sheet1";
		const ws = writer.addWorksheet(sheetName);
		let header: CellValue[] | null = null;
		let colIndexes: number[] = [];
		let i = 0;

		for await (const row of sheet) {
			i++;
			if (i <= skipRows) continue;
			const values = Array.isArray(row.values) ? row.values : [];

			if (!header) {
				const headerRow = values;
				header = headerRow;
				colIndexes = headerRow.map((h, idx) => (columnFilter(String(h ?? ""), idx) ? idx : -1)).filter((idx) => idx !== -1);
				ws.addRow(colIndexes.map((idx) => headerRow[idx])).commit();
				continue;
			}

			const filtered = colIndexes.map((idx) => values[idx]);
			if (rowFilter(filtered.map((cell) => String(cell ?? "")))) {
				ws.addRow(filtered).commit();
			}
		}
		ws.commit();
	}
	await writer.commit();
}
