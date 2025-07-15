import ExcelJS from "exceljs";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import { iterateToArray } from "microsoft-graph/iteration";
import NotFoundError from "microsoft-graph/NotFoundError";
import picomatch from "picomatch";
import type { Cell } from "../models/Cell.ts";
import type { Handle } from "../models/Handle.ts";
import type { CellRef, RangeRef } from "../models/Reference.ts";
import type { RowWrite } from "../models/Row.ts";
import type { DeleteShift, InsertShift } from "../models/Shift.ts";
import type { WorksheetName } from "../models/Worksheet.ts";
import { normalizeCellWrite } from "../services/cell.ts";
import { fromExcelCell, getWorksheetByName, updateExcelCell } from "../services/excelJs.ts";
import { parseCellReference, parseRangeReference } from "../services/reference.ts";
import { getLatestRevisionFilePath, getNextRevisionFilePath } from "../services/workingFolder.ts";

type TransactContext = (operations: TransactOperations) => Promise<void>;

type TransactOperations = {
	listWorksheets: () => WorksheetName[];
	tryFindWorksheet: (search: string) => WorksheetName | null;
	findWorksheet: (search: string) => WorksheetName;

	readCells: (range: RangeRef) => Cell[][];
	updateCells: (origin: CellRef, cells: Iterable<RowWrite> | AsyncIterable<RowWrite>) => Promise<void>;
	insertCells: (origin: CellRef, shift: InsertShift, cells: Iterable<RowWrite> | AsyncIterable<RowWrite>) => Promise<void>;
	deleteCells: (range: RangeRef, shift: DeleteShift) => void;
};

export default async function transactWorkbook(handle: Handle, context: TransactContext): Promise<void> {
	const file = await getLatestRevisionFilePath(handle.id);
	let isDirty = false;
	const workbook = new ExcelJS.Workbook();

	await workbook.xlsx.readFile(file);
	// TODO: Named ranges
	// TODO: Merging cells
	// TODO: Formulas in values

	const listWorksheets = () => workbook.worksheets.map((ws) => ws.name) as WorksheetName[];
	const tryFindWorksheet = (search: string): WorksheetName | null => {
		const matcher = picomatch(search, { nocase: true });
		const found = workbook.worksheets.find((ws) => matcher(ws.name));
		if (!found) return null;
		return found.name as WorksheetName;
	};
	const findWorksheet = (search: string): WorksheetName => {
		const worksheet = tryFindWorksheet(search);
		if (!worksheet) throw new NotFoundError(`Worksheet not found for search: ${search}`);
		return worksheet;
	};
	const readCells = (range: RangeRef): Cell[][] => {
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
	};
	const updateCells = async (origin: CellRef, cells: Iterable<RowWrite> | AsyncIterable<RowWrite>): Promise<void> => {
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
		isDirty = true;
	};

	const insertCells = async (origin: CellRef, shift: InsertShift, cells: Iterable<RowWrite> | AsyncIterable<RowWrite>): Promise<void> => {
		const [worksheetName, col, row] = parseCellReference(origin);
		const worksheet = getWorksheetByName(workbook, worksheetName);

		const rows = await iterateToArray(cells, (row) => row.map(normalizeCellWrite));

		if (shift === "Down") {
			worksheet.spliceRows(row, 0, ...rows.map((r) => r.map((cell) => cell.value)));
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
				const insertVals = Array.from({ length: maxCols }, (_, cIdx) => inputRow[cIdx]?.value);
				excelRow.splice(col, 0, ...insertVals);
				for (let c = 0; c < maxCols; c++) {
					updateExcelCell(excelRow.getCell(col + c), inputRow[c] ?? {});
				}
			});
		} else {
			throw new InvalidArgumentError(`Unsupported shift: ${shift}`);
		}
		
		isDirty = true;
	};

	const deleteCells = (range: RangeRef, shift: DeleteShift) => {
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
		isDirty = true;
	};

	const operations: TransactOperations = {
		listWorksheets,
		tryFindWorksheet,
		findWorksheet,
		readCells,
		updateCells,
		insertCells,
		deleteCells,
	};

	await context(operations);

	if (isDirty) {
		const nextFile = await getNextRevisionFilePath(handle.id);
		await workbook.xlsx.writeFile(nextFile);
	}
}
