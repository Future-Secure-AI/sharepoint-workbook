/**
 * Read a rectangular block of cells from a worksheet.
 * @module readCells
 * @category Tasks
 */

import type { Worksheet } from "aspose.cells.node";
import type { Cell } from "../models/Cell.ts";
import type { RangeRef } from "../models/Reference.ts";
import { readCell } from "../services/cellReader.ts";
import { parseRangeReferenceExact } from "../services/reference.ts";

/**
 * Reads a rectangular block of cells from the worksheet.
 * @param {Worksheet} worksheet The worksheet to read from.
 * @param {RangeRef} range The range reference (e.g., "A1:B2") specifying the block to read.
 * @returns {Cell[][]} A 2D array of Cell objects representing the values in the specified range.
 */
export function readCells(worksheet: Worksheet, range: RangeRef): Cell[][] {
	const [ac, ar, bc, br] = parseRangeReferenceExact(range, worksheet);

	const cells: Cell[][] = [];
	for (let r = ar; r <= br; r++) {
		const row: Cell[] = [];
		for (let c = ac; c < bc; c++) {
			const cell = readCell(worksheet, r, c);
			row.push(cell);
		}
		cells.push(row);
	}
	return cells;
}
