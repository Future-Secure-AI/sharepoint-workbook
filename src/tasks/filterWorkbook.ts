/**
 * Filter out unwanted rows and columns from a workbook.
 * @module filterWorkbook
 * @category Tasks
 */

import ExcelJS from "exceljs";
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

	progress?: (rows: number) => void;
};

/**
 * Filter out unwanted rows and columns from a workbook. All styling is lost when filtering.
 * @param handle Workbook handle to filter.
 * @param filter Filter options to apply (skipRows, column, row).
 * @returns A promise that resolves when the filtering is complete.
 */

export default async function filterWorkbook(handle: Handle, filter: Filter): Promise<void> {
	const { skipRows = 0, columnFilter = () => true, rowFilter = () => true, progress = () => {} } = filter;
	const latestFile = await getLatestRevisionFilePath(handle.id);
	const nextFile = await getNextRevisionFilePath(handle.id);

	const reader = new ExcelJS.stream.xlsx.WorkbookReader(latestFile, {});
	const writer = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: nextFile });

	for await (const sheet of reader) {
		const ws = writer.addWorksheet((sheet as { name?: string }).name ?? "Sheet1");
		let colIndexes: number[] | null = null;
		let lastProgressTime = 0;

		let r = 0;
		for await (const row of sheet) {
			r++;
			if (r <= skipRows) continue;
			const values = Array.isArray(row.values) ? row.values : [];

			if (colIndexes === null) {
				colIndexes = values.map((h, idx) => (columnFilter(String(h ?? ""), idx) ? idx : -1)).filter((idx) => idx !== -1);
				const row = ws.addRow(colIndexes.map((idx) => values[idx]));
				row.commit();
			} else {
				const filtered = colIndexes.map((idx) => values[idx]);
				if (rowFilter(filtered.map((cell) => String(cell ?? "")))) {
					const row = ws.addRow(filtered);
					row.commit();
				}
			}

			if (Date.now() - lastProgressTime >= 1000) {
				progress(r);
				lastProgressTime = Date.now();
			}
		}
		progress(r);
		ws.commit();
	}
	await writer.commit();
}
