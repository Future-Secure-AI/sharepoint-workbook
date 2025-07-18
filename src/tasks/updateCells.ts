/**
 * Update a rectangular block of cells in a worksheet, starting at the given origin.
 * @module updateCells
 * @category Tasks
 */
import type { Worksheet } from "aspose.cells.node";
import type { Cell, CellValue } from "../models/Cell.ts";
import type { ColumnIndex } from "../models/Column.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";
import type { CellRef } from "../models/Reference.ts";
import type { RowIndex } from "../models/Row.ts";
import { writeCell } from "../services/cellWriter.ts";
import { parseCellRef } from "../services/reference.ts";

/**
 * Updates a rectangular block of cells in the worksheet, starting at the given origin.
 * @param {Worksheet} worksheet The worksheet to update.
 * @param {CellRef} origin The top-left cell reference (e.g., "A1") where the update begins.
 * @param {(CellValue | DeepPartial<Cell>)[][]} cells A 2D array of cell values or partial cell objects to write. All rows must have the same length.
 */
export function updateCells(worksheet: Worksheet, origin: CellRef, cells: (CellValue | DeepPartial<Cell> | undefined)[][]): void {
	const [c, r] = parseCellRef(origin);

	for (let rx = 0; rx < cells.length; rx++) {
		const row = cells[rx] || [];
		for (let cx = 0; cx < row.length; cx++) {
			const value = row[cx];
			if (value === undefined) continue;
			writeCell(worksheet, (r + rx) as RowIndex, (c + cx) as ColumnIndex, value);
		}
	}
}
