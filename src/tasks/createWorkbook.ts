/**
 * Create a new workbook, optionally with specified worksheets.
 * @module createWorkbook
 * @category Tasks
 */

import AsposeCells from "aspose.cells.node";
import type { Cell, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";
import type { Workbook } from "../models/Workbook.ts";
import { writeCell } from "../services/cellWriter.ts";

/**
 * Create a new workbook, optionally with specified worksheets.
 * @param {Record<string, Iterable<RowValues> | AsyncIterable<RowValues>>} worksheets An object whose keys are worksheet names and values are iterables or async iterables of row values.
 * @returns {Promise<Handle>} Handle referencing the newly created workbook.
 * @example
 * const handle = await createWorkbook({
 *   Sheet1: [
 *     [1, 2, 3],
 *     [4, 5, 6],
 *   ],
 *   Sheet2: [
 *     ["A", "B", "C"],
 *     ["D", "E", "F"],
 *   ],
 * });
 */
export default async function createWorkbook(worksheets?: Record<string, (CellValue | DeepPartial<Cell>)[][]>): Promise<Workbook> {
	const workbook = new AsposeCells.Workbook();
	workbook.worksheets.removeAt(0); // Remove the default empty worksheet

	if (worksheets) {
		for (const [name, rows] of Object.entries(worksheets)) {
			const worksheet = workbook.worksheets.add(name);
			let r = 0;
			for (const row of rows) {
				let c = 0;
				for (const cellOrValue of row) {
					writeCell(worksheet, r, c, cellOrValue);
					c++;
				}
				r++;
			}
		}
	}

	return workbook;
}
