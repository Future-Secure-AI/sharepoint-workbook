/**
 * Insert a rectangular block of cells into a worksheet, shifting existing cells down or right.
 * @module insertCells
 * @category Tasks
 */

import type { Worksheet } from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { Cell, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";
import type { CellRef } from "../models/Reference.ts";
import type { InsertShift } from "../models/Shift.ts";
import { ensureRectangularArray } from "../services/rectangularArray.ts";
import { parseCellRef } from "../services/reference.ts";
import { updateCells } from "./updateCells.ts";

/**
 * Inserts a rectangular block of cells into the worksheet, shifting existing cells either down or right.
 * @param {Worksheet} worksheet The worksheet to modify.
 * @param {CellRef} origin The top-left cell reference (e.g., "A1") where the insertion begins.
 * @param {InsertShift} shift The direction to shift existing cells: "Down" to insert rows, "Right" to insert columns.
 * @param {(CellValue | DeepPartial<Cell>)[][]} cells A 2D rectangular array of cell values or partial cell objects to insert. All rows must have the same length.
 * @throws {InvalidArgumentError} If rows in `cells` have different lengths, or if `shift` is not "Down" or "Right".
 */
export function insertCells(worksheet: Worksheet, origin: CellRef, shift: InsertShift, cells: (CellValue | DeepPartial<Cell>)[][]): void {
	const [c, r] = parseCellRef(origin);

	ensureRectangularArray(cells);

	if (shift === "Down") {
		const count = cells.length;
		worksheet.cells.insertRows(r, count);
	} else if (shift === "Right") {
		const count = cells[0]?.length || 0;
		worksheet.cells.insertColumns(c, count);
	} else {
		throw new InvalidArgumentError(`Unsupported shift: ${shift}`);
	}

	updateCells(worksheet, origin, cells);
}
