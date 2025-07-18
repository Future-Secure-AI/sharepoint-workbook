/**
 * Deletes a given set of columns or rows from a worksheet.
 * @module deleteCells
 * @category Tasks
 */

import { type Worksheet, ShiftType } from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/dist/cjs/errors/InvalidArgumentError";
import type { ColumnRef, RowRef } from "../models/Reference.ts";
import { parseRef } from "../services/reference.ts";

export type ColumnOrRowRangeRef = `${ColumnRef | ""}:${ColumnRef | ""}` | `${RowRef | ""}:${RowRef | ""}` | [start: ColumnRef | null, end: ColumnRef | null] | [start: RowRef | null, end: RowRef | null];

/**
 * Deletes a given set of columns or rows from a worksheet. Adjacent cells will be shifted up or left.
 * @param worksheet The worksheet to modify.
 * @param range The range reference (e.g., "A:C" or "1:5") specifying the range to delete.
 * @throws {InvalidArgumentError} If shift is not "Up" or "Left".
 */

export function deleteCells(worksheet: Worksheet, range: ColumnOrRowRangeRef): void {
	const [ac, ar, bc, br] = parseRef(range);

	if (ac === null && bc === null && ar !== null && br !== null) {
		worksheet.cells.deleteRange(ar, 0, br, worksheet.cells.maxDataColumn, ShiftType.Up);
	} else if (ar === null && br === null && ac !== null && bc !== null) {
		worksheet.cells.deleteRange(0, ac, worksheet.cells.maxDataRow, bc, ShiftType.Left);
	} else {
		throw new InvalidArgumentError("Invalid range for deleteCells: must specify either a row or column range.");
	}
}
