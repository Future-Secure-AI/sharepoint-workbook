/**
 * Clear all values and formatting in a specified range of cells in a worksheet.
 * @module clearCells
 * @category Tasks
 */

import { ShiftType, type Worksheet } from "aspose.cells.node";
import type { Ref } from "../models/Reference.ts";
import { parseRefResolved } from "../services/reference.ts";

/**
 * Clear all values and formatting in a specified range of cells in a worksheet.
 * @param worksheet
 * @param range
 */
export function clearCells(worksheet: Worksheet, range: Ref): void {
	const [ac, ar, bc, br] = parseRefResolved(range, worksheet);
	worksheet.cells.deleteRange(ar, ac, br, bc, ShiftType.None);
}
