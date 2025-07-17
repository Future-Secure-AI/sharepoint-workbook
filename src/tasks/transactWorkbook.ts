import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import picomatch from "picomatch";
import type { Cell, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";
import type { Handle } from "../models/Handle.ts";
import type { CellRef, RangeRef } from "../models/Reference.ts";
import type { DeleteShift, InsertShift } from "../models/Shift.ts";
import type { WorksheetName } from "../models/Worksheet.ts";
import { parseCellReference, parseRangeReference } from "../services/reference.ts";

import type AsposeCells from "aspose.cells.node";
import { applyCell } from "../services/cell.ts";
// TODO: Named ranges, tables, pivot tables
// TODO: Merging cells

export function listWorksheets({ workbook }: Handle): WorksheetName[] {
	const result: WorksheetName[] = [];
	for (let i = 0; i < workbook.worksheets.count; i++) {
		const ws = workbook.worksheets.get(i);
		result.push(ws.name as WorksheetName);
	}
	return result;
}

export function tryFindWorksheet({ workbook }: Handle, search: string): WorksheetName | null {
	const matcher = picomatch(search, { nocase: true });

	for (let i = 0; i < workbook.worksheets.count; i++) {
		const ws = workbook.worksheets.get(i);
		if (matcher(ws.name)) {
			return ws.name as WorksheetName;
		}
	}
	return null;
}

function getWorksheetByName({ workbook }: Handle, name: WorksheetName): AsposeCells.Worksheet {
	const worksheet = workbook.worksheets.get(name);
	if (!worksheet) {
		throw new InvalidArgumentError(`Worksheet not found: ${name}`);
	}
	return worksheet;
}

// TODO: Following methods take a worksheet handle, and References are adjusted accordingly
export function insertCells({ workbook }: Handle, origin: CellRef, shift: InsertShift, cells: (CellValue | DeepPartial<Cell>)[][]): void {
	const [worksheetName, col, row] = parseCellReference(origin);
	const worksheet = getWorksheetByName({ workbook }, worksheetName);

	if (shift === "Down") {
		// TODO: Insert shifting down
	} else if (shift === "Right") {
		// TODO: Insert shifting right
	} else {
		throw new InvalidArgumentError(`Unsupported shift: ${shift}`);
	}
}

export function readCells({ workbook }: Handle, range: RangeRef): Cell[][] {
	let [worksheetName, startCol, startRow, endCol, endRow] = parseRangeReference(range);
	const worksheet = getWorksheetByName({ workbook }, worksheetName);

	startCol = startCol ?? worksheet.cells.minDataColumn;
	startRow = startRow ?? worksheet.cells.minDataRow;
	endCol = endCol ?? worksheet.cells.maxDataColumn;
	endRow = endRow ?? worksheet.cells.maxDataRow;

	// TODO: Read cells
}

export function updateCells({ workbook }: Handle, origin: CellRef, cells: (CellValue | DeepPartial<Cell>)[][]): void {
	const [worksheetName, col, row] = parseCellReference(origin);
	const worksheet = getWorksheetByName({ workbook }, worksheetName);

	// TODO: Update cells
}

export function deleteCells({ workbook }: Handle, range: RangeRef, shift: DeleteShift): void {
	let [worksheetName, startCol, startRow, endCol, endRow] = parseRangeReference(range);
	const worksheet = getWorksheetByName({ workbook }, worksheetName);

	startCol = startCol ?? worksheet.cells.minDataColumn;
	startRow = startRow ?? worksheet.cells.minDataRow;
	endCol = endCol ?? worksheet.cells.maxDataColumn;
	endRow = endRow ?? worksheet.cells.maxDataRow;

	if (shift === "Up") {
		// TODO: Delete shifting up
	} else if (shift === "Left") {
		// TODO: Delete shifting left
	} else {
		throw new InvalidArgumentError(`Unsupported shift: ${shift}`);
	}
}

export function updateEachCell({ workbook }: Handle, range: RangeRef, write: CellValue | DeepPartial<Cell>): void {
	let [worksheetName, startCol, startRow, endCol, endRow] = parseRangeReference(range);
	const worksheet = getWorksheetByName({ workbook }, worksheetName);

	startCol = startCol ?? worksheet.cells.minDataColumn;
	startRow = startRow ?? worksheet.cells.minDataRow;
	endCol = endCol ?? worksheet.cells.maxDataColumn;
	endRow = endRow ?? worksheet.cells.maxDataRow;

	for (let r = startRow; r <= endRow; r++) {
		for (let c = startCol; c <= endCol; c++) {
			applyCell(worksheet, r, c, write);
		}
	}
}
