/**
 * Update every cell in a rectangular range to the same value or partial cell object.
 * @module updateEachCell
 * @category Tasks
 */
import type { Worksheet } from "aspose.cells.node";
import type { Cell, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";
import type { RangeRef } from "../models/Reference.ts";
import { writeCell } from "../services/cellWriter.ts";
import { parseRangeReferenceExact } from "../services/reference.ts";

/**
 * Updates every cell in the specified rectangular range to the given value or partial cell object.
 * @param {Worksheet} worksheet The worksheet to update.
 * @param {RangeRef} range The range reference (e.g., "A1:B2") specifying the block to update.
 * @param {CellValue | DeepPartial<Cell>} write The value or partial cell object to write to each cell in the range.
 */
export function updateEachCell(worksheet: Worksheet, range: RangeRef, write: CellValue | DeepPartial<Cell>): void {
	const [ac, ar, bc, br] = parseRangeReferenceExact(range, worksheet);

	for (let r = ar; r <= br; r++) {
		for (let c = ac; c <= bc; c++) {
			writeCell(worksheet, r, c, write);
		}
	}
}
