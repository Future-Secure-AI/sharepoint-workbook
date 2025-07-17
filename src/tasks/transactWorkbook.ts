import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import { iterateToArray } from "microsoft-graph/iteration";
import NotFoundError from "microsoft-graph/NotFoundError";
import picomatch from "picomatch";
import type { Cell, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";
import type { Handle } from "../models/Handle.ts";
import type { CellRef, RangeRef } from "../models/Reference.ts";
import type { DeleteShift, InsertShift } from "../models/Shift.ts";
import type { WorksheetName } from "../models/Worksheet.ts";
import { parseCellReference, parseRangeReference } from "../services/reference.ts";

// TODO: Named ranges, tables, pivot tables
// TODO: Merging cells

export function listWorksheets({ workbook }: Handle): WorksheetName[] {
	return workbook.worksheets.map((ws) => ws.name) as WorksheetName[];
}

export function tryFindWorksheet({ workbook }: Handle, search: string): WorksheetName | null {
	const matcher = picomatch(search, { nocase: true });
	const found = workbook.worksheets.find((ws) => matcher(ws.name));
	if (!found) return null;
	return found.name as WorksheetName;
}

export function findWorksheet({ workbook }: Handle, search: string): WorksheetName {
	const worksheet = tryFindWorksheet(search);
	if (!worksheet) throw new NotFoundError(`Worksheet not found for search: ${search}`);
	return worksheet;
}

export function readCells({ workbook }: Handle, range: RangeRef): Cell[][] {
	let [worksheetName, startCol, startRow, endCol, endRow] = parseRangeReference(range);
	const worksheet = getWorksheetByName(workbook, worksheetName);

	startCol = startCol ?? 1;
	startRow = startRow ?? 1;
	endCol = endCol ?? worksheet.columnCount;
	endRow = endRow ?? worksheet.rowCount;

	const outputRows: Cell[][] = [];
	for (let r = startRow; r <= endRow; r = r + 1) {
		const inputRow = worksheet.getRow(r);
		const outputRow: Cell[] = [];
		for (let c = startCol; c <= endCol; c++) {
			const inputCell = inputRow.getCell(c);
			const outputCell = fromExcelCell(inputCell);
			outputRow.push(outputCell);
		}
		outputRows.push(outputRow);
	}
	return outputRows;
}

export function updateCells({ workbook }: Handle, origin: CellRef, cells: (CellValue | DeepPartial<Cell>)[][]): void {
	const [worksheetName, col, row] = parseCellReference(origin);
	const worksheet = getWorksheetByName(workbook, worksheetName);

	let r = 0;
	for await (const inputRow of cells) {
		const excelRow = worksheet.getRow(row + r);

		let c = 0;
		for (const inputCell of inputRow) {
			const write = normalizeCellWrite(inputCell);
			const excelCell = excelRow.getCell(col + c);
			updateExcelCell(excelCell, write);
			c++;
		}
		r++;
	}
}

export async function insertCells({ workbook }: Handle, origin: CellRef, shift: InsertShift, cells: (CellValue | DeepPartial<Cell>)[][]): void {
	const [worksheetName, col, row] = parseCellReference(origin);
	const worksheet = getWorksheetByName(workbook, worksheetName);

	const rows = await iterateToArray(cells, (row) => row.map(normalizeCellWrite));

	if (shift === "Down") {
		worksheet.spliceRows(row, 0, ...rows.map((r) => r.map((cell) => toExcelValue(cell.value))));
		rows.forEach((inputRow, r) => {
			const excelRow = worksheet.getRow(row + r);

			inputRow.forEach((write, c) => {
				const excelCell = excelRow.getCell(col + c);
				updateExcelCell(excelCell, write);
			});
		});
	} else if (shift === "Right") {
		const maxCols = Math.max(0, ...rows.map((r) => r.length));
		rows.forEach((inputRow, r) => {
			const excelRow = worksheet.getRow(row + r);
			const insertVals = Array.from({ length: maxCols }, (_, cIdx) => toExcelValue(inputRow[cIdx]?.value));
			excelRow.splice(col, 0, ...insertVals);
			for (let c = 0; c < maxCols; c++) {
				updateExcelCell(excelRow.getCell(col + c), inputRow[c] ?? {});
			}
		});
	} else {
		throw new InvalidArgumentError(`Unsupported shift: ${shift}`);
	}
}

export function deleteCells({ workbook }: Handle, range: RangeRef, shift: DeleteShift): void {
	let [worksheetName, startCol, startRow, endCol, endRow] = parseRangeReference(range);
	const worksheet = getWorksheetByName(workbook, worksheetName);

	startCol = startCol ?? 1;
	startRow = startRow ?? 1;
	endCol = endCol ?? worksheet.columnCount;
	endRow = endRow ?? worksheet.rowCount;

	if (shift === "Up") {
		const numRows = endRow - startRow + 1;
		worksheet.spliceRows(startRow, numRows);
	} else if (shift === "Left") {
		const numCols = endCol - startCol + 1;
		for (let r = 1; r <= worksheet.rowCount; r++) {
			const row = worksheet.getRow(r);
			row.splice(startCol, numCols);
			row.commit();
		}
	} else {
		throw new InvalidArgumentError(`Unsupported shift: ${shift}`);
	}

	worksheet.commit();
}

export function updateEachCell({ workbook }: Handle, range: RangeRef, write: CellWrite): void {
	let [worksheetName, startCol, startRow, endCol, endRow] = parseRangeReference(range);
	const worksheet = getWorksheetByName(workbook, worksheetName);

	startCol = startCol ?? 1;
	startRow = startRow ?? 1;
	endCol = endCol ?? worksheet.columnCount;
	endRow = endRow ?? worksheet.rowCount;

	for (let r = startRow; r <= endRow; r++) {
		const excelRow = worksheet.getRow(r);
		for (let c = startCol; c <= endCol; c++) {
			const excelCell = excelRow.getCell(c);
			updateExcelCell(excelCell, write);
		}
	}
}
