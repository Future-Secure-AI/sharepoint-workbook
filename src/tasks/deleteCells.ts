import { type Worksheet, ShiftType } from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { RangeRef } from "../models/Reference.ts";
import type { DeleteShift } from "../models/Shift.ts";
import { parseRangeReferenceExact } from "../services/reference.ts";

/**
 * Deletes a rectangular block of cells from the worksheet, shifting remaining cells up or left.
 * @param worksheet The worksheet to modify.
 * @param range The range reference (e.g., "A1:B2") specifying the block to delete.
 * @param shift The direction to shift remaining cells: "Up" or "Left".
 * @throws {InvalidArgumentError} If shift is not "Up" or "Left".
 */

export function deleteCells(worksheet: Worksheet, range: RangeRef, shift: DeleteShift): void {
	const [ac, ar, bc, br] = parseRangeReferenceExact(range, worksheet);

	if (shift === "Up") {
		worksheet.cells.deleteRange(ac, ar, bc, br, ShiftType.Up); // 0 = ShiftUp
	} else if (shift === "Left") {
		worksheet.cells.deleteRange(ac, ar, bc, br, ShiftType.Left); // 1 = ShiftLeft
	} else {
		throw new InvalidArgumentError(`Unsupported shift: ${shift}`);
	}
}
