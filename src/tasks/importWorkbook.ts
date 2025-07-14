/**
 * Imports worksheet content as a new open workbook.
 * @module importWorkbook
 * @category Tasks
 */

import ExcelJS from "exceljs";
import { createWriteStream } from "node:fs";
import type { Handle } from "../models/Handle.ts";
import type { WriteWorksheet } from "../models/Worksheet.ts";
import { appendRow } from "../services/excelJs.ts";
import { createHandleId, getNextRevisionFilePath } from "../services/workingFolder.ts";

/**
 * Imports worksheet content as a new open workbook.
 * @param {Iterable<WriteWorksheet> | AsyncIterable<WriteWorksheet>} worksheets Worksheet data to import.
 * @returns {Promise<Handle>} Handle referencing the newly created workbook.
 */
export default async function importWorkbook(worksheets: Iterable<WriteWorksheet> | AsyncIterable<WriteWorksheet>): Promise<Handle> {
	const id = createHandleId();
	const file = await getNextRevisionFilePath(id);

	const rawStream = createWriteStream(file);
	const xls = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: rawStream });
	for await (const { name, rows } of worksheets) {
		const worksheet = xls.addWorksheet(name);
		for await (const row of rows) {
			appendRow(worksheet, row);
		}
		worksheet.commit(); // Ensure worksheet data is flushed
	}
	await xls.commit();

	return {
		id,
	};
}
