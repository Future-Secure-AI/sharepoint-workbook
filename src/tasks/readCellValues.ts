import type { Worksheet } from "aspose.cells.node";
import type { CellValue } from "../models/Cell.ts";
import type { RangeRef } from "../models/Reference.ts";
import { readCellValue } from "../services/cell.ts";
import { parseRangeReferenceExact } from "../services/reference.ts";

/**
 * Reads a rectangular block of cell values from the worksheet. No styles are included.
 * @param worksheet The worksheet to read from.
 * @param range The range reference (e.g., "A1:B2") specifying the block to read.
 * @returns A 2D array of CellValue objects representing the values in the specified range.
 */

export function readCellValues(worksheet: Worksheet, range: RangeRef): CellValue[][] {
	const [ac, ar, bc, br] = parseRangeReferenceExact(range, worksheet);

	const cells: CellValue[][] = [];
	for (let r = ar; r <= br; r++) {
		const row: CellValue[] = [];
		for (let c = ac; c < bc; c++) {
			const cell = readCellValue(worksheet, r, c);
			row.push(cell);
		}
		cells.push(row);
	}
	return cells;
}
