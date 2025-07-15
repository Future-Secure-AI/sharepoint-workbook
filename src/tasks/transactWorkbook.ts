import type { CellValue, Column, Row } from "exceljs";
import ExcelJS from "exceljs";
import NotFoundError from "microsoft-graph/NotFoundError";
import picomatch from "picomatch";
import type { Cell, CellWrite } from "../models/Cell.ts";
import type { Handle } from "../models/Handle.ts";
import type { CellRef, RangeRef } from "../models/Reference.ts";
import type { DeleteShift, InsertShift } from "../models/Shift.ts";
import type { WorksheetName } from "../models/Worksheet.ts";
import { getCell } from "../services/excelJs.ts";
import { parseRangeReference } from "../services/reference.ts";
import { getLatestRevisionFilePath, getNextRevisionFilePath } from "../services/workingFolder.ts";

const pattern = /^(?<col>[A-Z]{0,3})(?<row>[0-9]{0,7})$/;
type TransactContext = (operations: TransactOperations) => void;

type TransactOperations = {
	listWorksheets: () => WorksheetName[];
	tryFindWorksheet: (search: string) => WorksheetName | null;
	findWorksheet: (search: string) => WorksheetName;

	readCells: (range: RangeRef) => Cell[][];
	updateCells: (origin: CellRef, cells: (CellValue | CellWrite)[][]) => void;
	insertCells: (origin: CellRef, shift: InsertShift, cells: (CellValue | CellWrite)[][]) => void;
	deleteCells: (range: RangeRef, shift: DeleteShift) => void;
};

export default async function transactWorkbook(handle: Handle, context: TransactContext): Promise<void> {
	const file = await getLatestRevisionFilePath(handle.id);
	const isDirty = false;
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.readFile(file);
	// TODO: Named ranges

	const listWorksheets = () => workbook.worksheets.map((ws) => ws.name) as WorksheetName[];
	const tryFindWorksheet = (search: string): WorksheetName | null => {
		const matcher = picomatch(search, { nocase: true });
		const found = workbook.worksheets.find((ws) => matcher(ws.name));
		if (!found) return null;
		return found.name as WorksheetName;
	};
	const findWorksheet = (search: string): WorksheetName => {
		const worksheet = tryFindWorksheet(search);
		if (!worksheet) {
			throw new NotFoundError(`Worksheet not found for search: ${search}`);
		}
		return worksheet;
	};
	const readCells = (range: RangeRef): Cell[][] => {
		let [worksheetName, startCol, startRow, endCol, endRow] = parseRangeReference(range);
		const worksheet = workbook.getWorksheet(worksheetName);
		if (!worksheet) {
			throw new NotFoundError(`Worksheet not found: ${worksheetName}`);
		}

		startCol = startCol ?? 1;
		startRow = startRow ?? 1;
		endCol = endCol ?? worksheet.columnCount;
		endRow = endRow ?? worksheet.rowCount;

		const output: Cell[][] = [];
		for (let r = startRow; r <= endRow; r = r + 1) {
			const rowCells: Cell[] = [];
			for (let c = startCol; c <= endCol; c++) {
				const cell = getCell(worksheet, r, c);

				rowCells.push(cell);
			}
			output.push(rowCells);
		}
		return output;
	};

	const operations: TransactOperations = {
		listWorksheets,
		tryFindWorksheet,
		findWorksheet,
		readCells,
	};

	context(operations);

	if (isDirty) {
		const nextFile = await getNextRevisionFilePath(handle.id);
		await workbook.xlsx.writeFile(nextFile);
	}
}

const h = {} as Handle;

await transactWorkbook(h, ({ findWorksheet, insertCells, deleteCells }) => {
	const worksheetName = findWorksheet("*");

	insertCells([worksheetName, "A1"], "Down", [
		["A1", "B1", "C1"],
		["A2", "B2", "C2"],
	]);

	deleteCells([worksheetName, 1, 5], "Up");
});

function decomposeRef(ref: Row | Column | Cell): { col: Column | null; row: Row | null } {
	if (typeof ref === "number") {
		return {
			col: null,
			row: `${ref}` as unknown as Row,
		};
	}

	const match = ref.toString().match(pattern);
	if (!match || !match.groups) {
		throw new Error(`Invalid reference format: ${ref}`);
	}

	return {
		// biome-ignore lint/complexity/useLiteralKeys: Not possible with RegExp groups
		col: match.groups["col"] as unknown as Column,
		// biome-ignore lint/complexity/useLiteralKeys: Not possible with RegExp groups
		row: match.groups["row"] as unknown as Row,
	};
}
