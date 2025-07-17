/**
 * Imports worksheet content as a new open workbook.
 * @module importWorkbook
 * @category Tasks
 */

import AsposeCells from "aspose.cells.node";
import type { Cell, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";
import type { Handle } from "../models/Handle.ts";
import { applyCell } from "../services/cell.ts";

/**
 * Imports worksheet content as a new open workbook.
 * @param {Record<string, Iterable<RowValues> | AsyncIterable<RowValues>>} worksheets An object whose keys are worksheet names and values are iterables or async iterables of row values.
 * @returns {Promise<Handle>} Handle referencing the newly created workbook.
 * @example
 * const handle = await importWorkbook({
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
export default async function importWorkbook(worksheets: Record<string, (CellValue | DeepPartial<Cell>)[][]>): Promise<Handle> {
	const workbook = new AsposeCells.Workbook();
	workbook.worksheets.removeAt(0); // Remove the default empty worksheet

	for (const [name, rows] of Object.entries(worksheets)) {
		const worksheet = workbook.worksheets.add(name);
		let r = 0;
		for (const row of rows) {
			let c = 0;
			for (const cellOrValue of row) {
				applyCell(workbook, worksheet, r, c, cellOrValue);
				c++;
			}
			r++;
		}
	}

	return {
		workbook,
	};
}
