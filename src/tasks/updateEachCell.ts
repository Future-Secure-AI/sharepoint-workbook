/**
 * Applies an update to every cell in the specified range of a worksheet.
 * @module updateEachCell
 * @category Tasks
 */
import type { Worksheet } from "aspose.cells.node";
import type { Cell, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";
import type { Ref } from "../models/Reference.ts";
import { writeCell } from "../services/cellWriter.ts";
import { parseRefResolved } from "../services/reference.ts";

/**
 * Applies an update to every cell in the specified range of a worksheet.
 * @param {Worksheet} worksheet The worksheet to update.
 * @param {Ref} range The range reference (e.g., "A1:B2") specifying the block to update.
 * @param {CellValue | DeepPartial<Cell>} write The value or partial cell object to write to each cell in the range.
 * @example
 * // Updates every cell in the range A1:B2 to have a value of 42
 * updateEachCell(worksheet, "A1:B2", 42);
 *
 * // Updates every cell in the first row to be bold
 * updateEachCell(worksheet, "1", { fontBold: true });
 */
export default function updateEachCell(worksheet: Worksheet, range: Ref, write: CellValue | DeepPartial<Cell>): void {
	const [ac, ar, bc, br] = parseRefResolved(range, worksheet);

	for (let r = ar; r <= br; r++) {
		for (let c = ac; c <= bc; c++) {
			writeCell(worksheet, r, c, write);
		}
	}
}
